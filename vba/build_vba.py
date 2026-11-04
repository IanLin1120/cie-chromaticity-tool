"""Build a source-only vbaProject.bin (MS-OVBA) inside a Compound File (MS-CFB v3).

Per MS-OVBA, the _VBA_PROJECT version is written as 0xFFFF and module performance caches are
omitted (MODULEOFFSET = 0), so the host (Excel) compiles the project from the stored source.
"""
import re, struct, random, sys, uuid

CODEPAGE = 1252

# ----------------------------------------------------------------- MS-OVBA compression (2.4.1)
def ovba_compress(data: bytes) -> bytes:
    out = bytearray(b'\x01')
    pos = 0
    while pos < len(data):
        chunk = data[pos:pos + 4096]
        comp = _compress_chunk(chunk)
        if len(comp) > 4096 and len(chunk) == 4096:          # store raw
            out += struct.pack('<H', 0x3000 | (4096 + 2 - 3)) + chunk
        else:
            out += struct.pack('<H', 0xB000 | (len(comp) + 2 - 3)) + comp
        pos += 4096
    return bytes(out)

def _copy_params(diff):
    bit = max((diff - 1).bit_length(), 4)
    length_mask = 0xFFFF >> bit
    return bit, length_mask, length_mask + 3

def _compress_chunk(ch: bytes) -> bytes:
    out = bytearray()
    i, n = 0, len(ch)
    while i < n:
        flag_pos = len(out); out.append(0); flag = 0
        for b in range(8):
            if i >= n: break
            bit, length_mask, max_len = _copy_params(i) if i > 0 else (4, 0xFFF, 0xFFF + 3)
            best_len, best_off = 0, 0
            if i > 0:
                start = max(0, i - (1 << bit))
                for j in range(i - 1, start - 1, -1):
                    L = 0
                    while i + L < n and L < max_len and ch[j + L] == ch[i + L]:
                        L += 1
                    if L > best_len:
                        best_len, best_off = L, i - j
                        if L == max_len: break
            if best_len >= 3:
                token = ((best_off - 1) << (16 - bit)) | (best_len - 3)
                out += struct.pack('<H', token); flag |= (1 << b); i += best_len
            else:
                out.append(ch[i]); i += 1
        out[flag_pos] = flag
    return bytes(out)

# ----------------------------------------------------------------- dir stream (2.3.4.2)
def rec(rid, payload=b''):
    return struct.pack('<HI', rid, len(payload)) + payload

def mbcs(s): return s.encode('cp1252')
def u16(s): return s.encode('utf-16-le')

def dir_stream(modules):
    d = bytearray()
    d += rec(0x0001, struct.pack('<I', 1))                 # SYSKIND win32
    d += rec(0x0002, struct.pack('<I', 0x409))             # LCID
    d += rec(0x0014, struct.pack('<I', 0x409))             # LCIDINVOKE
    d += rec(0x0003, struct.pack('<H', CODEPAGE))          # CODEPAGE
    d += rec(0x0004, mbcs('VBAProject'))                   # NAME
    d += rec(0x0005, b'') + rec(0x0040, b'')               # DOCSTRING (+unicode)
    d += rec(0x0006, b'') + rec(0x003D, b'')               # HELPFILEPATH 1/2
    d += rec(0x0007, struct.pack('<I', 0))                 # HELPCONTEXT
    d += rec(0x0008, struct.pack('<I', 0))                 # LIBFLAGS
    d += struct.pack('<HIIH', 0x0009, 4, 1, 0)             # VERSION (Reserved=4, major, minor)
    d += rec(0x000C, b'') + rec(0x003C, b'')               # CONSTANTS (+unicode)
    # no references
    d += rec(0x000F, struct.pack('<H', len(modules)))      # MODULES count
    d += rec(0x0013, struct.pack('<H', 0xFFFF))            # PROJECTCOOKIE
    for name, kind, _ in modules:
        d += rec(0x0019, mbcs(name))
        d += rec(0x0047, u16(name))
        d += rec(0x001A, mbcs(name)) + rec(0x0032, u16(name))
        d += rec(0x001C, b'') + rec(0x0048, b'')
        d += rec(0x0031, struct.pack('<I', 0))             # MODULEOFFSET = 0 (source only)
        d += rec(0x001E, struct.pack('<I', 0))
        d += rec(0x002C, struct.pack('<H', 0xFFFF))
        d += rec(0x0021 if kind == 'std' else 0x0022, b'')
        d += rec(0x002B, b'')                              # module terminator
    d += rec(0x0010, b'')                                  # dir terminator
    return bytes(d)

# ----------------------------------------------------------------- PROJECT stream data encryption (2.4.3)
def ovba_encrypt(data: bytes, project_id: str) -> str:
    seed = random.randint(0, 255)
    version = 2
    key = sum(project_id.encode('cp1252')) & 0xFF
    out = bytearray([seed, seed ^ version, seed ^ key])
    unenc1, enc1, enc2 = key, seed ^ key, seed ^ version
    for _ in range((seed & 6) >> 1):
        t = random.randint(0, 255)
        be = t ^ ((enc2 + unenc1) & 0xFF)
        out.append(be); enc2, enc1, unenc1 = enc1, be, t
    for byte in struct.pack('<I', len(data)) + data:
        be = byte ^ ((enc2 + unenc1) & 0xFF)
        out.append(be); enc2, enc1, unenc1 = enc1, be, byte
    return out.hex().upper()

def ovba_decrypt(hexs: str, project_id: str) -> bytes:   # for self-check
    b = bytes.fromhex(hexs); seed, venc, kenc = b[0], b[1], b[2]
    key = seed ^ kenc; assert key == sum(project_id.encode('cp1252')) & 0xFF and (seed ^ venc) == 2
    unenc1, enc1, enc2 = key, kenc, venc; i = 3; outb = []
    for _ in range((seed & 6) >> 1):
        be = b[i]; t = be ^ ((enc2 + unenc1) & 0xFF); enc2, enc1, unenc1 = enc1, be, t; i += 1
    rest = []
    while i < len(b):
        be = b[i]; byte = be ^ ((enc2 + unenc1) & 0xFF); rest.append(byte); enc2, enc1, unenc1 = enc1, be, byte; i += 1
    ln = struct.unpack('<I', bytes(rest[:4]))[0]
    return bytes(rest[4:4 + ln])

def project_stream(modules):
    pid = '{' + str(uuid.uuid4()).upper() + '}'
    lines = [f'ID="{pid}"']
    for name, kind, _ in modules:
        lines.append(f'Document={name}/&H00000000' if kind == 'doc' else f'Module={name}')
    lines += ['Name="VBAProject"', 'HelpContextID="0"', 'VersionCompatible32="393222000"',
              f'CMG="{ovba_encrypt(struct.pack("<I", 0), pid)}"',
              f'DPB="{ovba_encrypt(b"\x00", pid)}"',
              f'GC="{ovba_encrypt(b"\xff", pid)}"',
              '', '[Host Extender Info]', '&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000', '',
              '[Workspace]'] + [f'{name}=0, 0, 0, 0, C' for name, _, _ in modules]
    for k, h in (('CMG', 4), ('DPB', 1), ('GC', 1)):
        v = [l for l in lines if l.startswith(k + '=')][0].split('"')[1]
        assert len(ovba_decrypt(v, pid)) == h
    return ('\r\n'.join(lines) + '\r\n').encode('cp1252')

def projectwm(modules):
    b = bytearray()
    for name, _, _ in modules:
        b += mbcs(name) + b'\x00' + u16(name) + b'\x00\x00'
    return bytes(b + b'\x00\x00')

# ----------------------------------------------------------------- MS-CFB v3 writer
FREESECT, ENDOFCHAIN, FATSECT, NOSTREAM = 0xFFFFFFFF, 0xFFFFFFFE, 0xFFFFFFFD, 0xFFFFFFFF

def cfb(tree):
    """tree: {'name': {...} for storage | bytes for stream} -> bytes"""
    entries = []    # dicts: name,type,data,children
    def add(name, node, typ):
        e = {'name': name, 'type': typ, 'data': None, 'kids': [], 'L': NOSTREAM, 'R': NOSTREAM, 'C': NOSTREAM, 'color': 1}
        entries.append(e); idx = len(entries) - 1
        if isinstance(node, dict):
            for k, v in node.items():
                e['kids'].append(add(k, v, 1 if isinstance(v, dict) else 2))
        else:
            e['data'] = node
        return idx
    add('Root Entry', tree, 5)

    key = lambda i: (len(entries[i]['name']), entries[i]['name'].upper())
    for e in entries:                                 # balanced BST of siblings, deepest level red
        kids = sorted(e['kids'], key=key)
        def build(lst, depth, out):
            if not lst: return NOSTREAM
            m = len(lst) // 2; node = lst[m]; out.append((node, depth))
            entries[node]['L'] = build(lst[:m], depth + 1, out)
            entries[node]['R'] = build(lst[m + 1:], depth + 1, out)
            return node
        info = []
        e['C'] = build(kids, 0, info)
        if info:
            maxd = max(d for _, d in info)
            full = all(len([1 for _, d in info if d == k]) == 2 ** k for k in range(maxd + 1))
            for node, d in info:
                entries[node]['color'] = 0 if (d == maxd and not full and maxd > 0) else 1

    # mini stream
    mini = bytearray(); minifat = []
    big = []
    for e in entries:
        if e['type'] != 2: continue
        data = e['data']
        if len(data) < 4096:
            start = len(mini) // 64; nsec = max(1, (len(data) + 63) // 64) if data else 0
            e['start'] = start if data else ENDOFCHAIN
            mini += data + b'\x00' * ((-len(data)) % 64)
            for k in range(nsec): minifat.append(start + k + 1 if k < nsec - 1 else ENDOFCHAIN)
        else:
            big.append(e)

    sectors = []      # list of 512-byte payloads (excluding FAT sectors, appended at start later)
    chains = {}
    def alloc(blob):
        blob = blob + b'\x00' * ((-len(blob)) % 512)
        first = len(sectors)
        for k in range(len(blob) // 512): sectors.append(blob[k * 512:(k + 1) * 512])
        return first, len(blob) // 512
    # directory
    dir_blob = bytearray()
    root_start = ENDOFCHAIN
    mini_first = mini_n = None
    if mini: mini_first, mini_n = alloc(bytes(mini))
    minifat_blob = b''.join(struct.pack('<I', v) for v in minifat)
    mf_first = mf_n = None
    if minifat: mf_first, mf_n = alloc(minifat_blob)
    for e in big:
        e['first'], e['n'] = alloc(e['data'])
    nent = len(entries) + ((-len(entries)) % 4)
    dir_first, dir_n = alloc(b'\x00' * (nent * 128))

    # FAT sectors needed (+ offset all others)
    total = len(sectors)
    nfat = 1
    while nfat * 128 < total + nfat: nfat += 1
    off = nfat
    fat = [FREESECT] * (nfat * 128)
    for k in range(nfat): fat[k] = FATSECT
    def chain(first, n):
        for k in range(n): fat[off + first + k] = off + first + k + 1 if k < n - 1 else ENDOFCHAIN
    if mini: chain(mini_first, mini_n)
    if minifat: chain(mf_first, mf_n)
    for e in big: chain(e['first'], e['n'])
    chain(dir_first, dir_n)

    # directory entries
    for i, e in enumerate(entries):
        nm = e['name'].encode('utf-16-le') + b'\x00\x00'
        b = nm + b'\x00' * (64 - len(nm))
        b += struct.pack('<HBB', len(nm), e['type'], e['color'])
        b += struct.pack('<III', e['L'], e['R'], e['C'])
        b += b'\x00' * 16 + struct.pack('<I', 0) + b'\x00' * 16
        if e['type'] == 5:
            b += struct.pack('<IQ', (off + mini_first) if mini else ENDOFCHAIN, len(mini))
        elif e['type'] == 2:
            if len(e['data']) >= 4096: b += struct.pack('<IQ', off + e['first'], len(e['data']))
            else: b += struct.pack('<IQ', e['start'], len(e['data']))
        else:
            b += struct.pack('<IQ', 0, 0)
        dir_blob += b
    for i in range(len(entries), nent):
        b = b'\x00' * 64 + struct.pack('<HBB', 0, 0, 0) + struct.pack('<III', NOSTREAM, NOSTREAM, NOSTREAM) + b'\x00' * 36 + struct.pack('<IQ', 0, 0)
        dir_blob += b
    for k in range(dir_n): sectors[dir_first + k] = bytes(dir_blob[k * 512:(k + 1) * 512])

    hdr = bytearray(b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1' + b'\x00' * 16)
    hdr += struct.pack('<HHHHH', 0x3E, 3, 0xFFFE, 9, 6) + b'\x00' * 6
    hdr += struct.pack('<IIIIIIIII', 0, nfat, off + dir_first, 0, 4096,
                       (off + mf_first) if minifat else ENDOFCHAIN, mf_n or 0, ENDOFCHAIN, 0)
    difat = [k for k in range(nfat)] + [FREESECT] * (109 - nfat)
    hdr += b''.join(struct.pack('<I', v) for v in difat)
    assert len(hdr) == 512
    fat_bytes = b''.join(struct.pack('<I', v) for v in fat)
    return bytes(hdr) + fat_bytes + b''.join(sectors)

# ----------------------------------------------------------------- source handling
def chrw_expr(text):
    parts, buf = [], ''
    for ch in text:
        if ord(ch) < 128 and ch != '"':
            buf += ch
        else:
            if buf: parts.append('"' + buf + '"'); buf = ''
            parts.append(f'ChrW(&H{ord(ch):X})' if ch != '"' else 'Chr(34)')
    if buf: parts.append('"' + buf + '"')
    return ' & '.join(parts) if parts else '""'

DOC_ATTR = {
    'ThisWorkbook': '0{00020819-0000-0000-C000-000000000046}',
    'Sheet': '0{00020820-0000-0000-C000-000000000046}',
}
def doc_module(name):
    base = DOC_ATTR['ThisWorkbook'] if name == 'ThisWorkbook' else DOC_ATTR['Sheet']
    return '\r\n'.join([f'Attribute VB_Name = "{name}"', f'Attribute VB_Base = "{base}"',
        'Attribute VB_GlobalNameSpace = False', 'Attribute VB_Creatable = False', 'Attribute VB_PredeclaredId = True',
        'Attribute VB_Exposed = True', 'Attribute VB_TemplateDerived = False', 'Attribute VB_Customizable = True']) + '\r\n'

def build(module_src_path, sheets=3, sheet1_code=None):
    src = open(module_src_path, encoding='utf-8').read()
    src = re.sub(r'\{\{(.*?)\}\}', lambda m: chrw_expr(m.group(1)), src)
    assert all(ord(c) < 128 for c in src), 'module source must be ASCII after ChrW conversion'
    src = src.replace('\r\n', '\n').replace('\n', '\r\n')
    modules = [('ThisWorkbook', 'doc', doc_module('ThisWorkbook'))]
    def extra(i):
        if i == 1 and sheet1_code:
            code = open(sheet1_code, encoding='utf-8').read().replace('\r\n', '\n').replace('\n', '\r\n')
            assert all(ord(c) < 128 for c in code)
            return code
        return ''
    modules += [(f'Sheet{i}', 'doc', doc_module(f'Sheet{i}') + extra(i)) for i in range(1, sheets + 1)]
    modules += [('Module1', 'std', src)]
    vba = {'_VBA_PROJECT': b'\xCC\x61\xFF\xFF\x00\x00\x00', 'dir': ovba_compress(dir_stream(modules))}
    for name, kind, text in modules:
        vba[name] = ovba_compress(text.encode('cp1252'))
    tree = {'VBA': vba, 'PROJECT': project_stream(modules), 'PROJECTwm': projectwm(modules)}
    return cfb(tree)

if __name__ == '__main__':
    out = sys.argv[2] if len(sys.argv) > 2 else 'vbaProject.bin'
    data = build(sys.argv[1], sheet1_code=sys.argv[3] if len(sys.argv) > 3 else None)
    open(out, 'wb').write(data); print(out, len(data), 'bytes')
