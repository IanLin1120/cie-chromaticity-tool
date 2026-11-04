import re, sys
def check(path, extra_globals=()):
    src=open(path,encoding='utf-8').read()
    src=re.sub(r'\{\{.*?\}\}','""',src)
    lines=[]
    for ln in src.split('\n'):
        ln=re.sub(r'"[^"]*"','""',ln); ln=re.sub(r"'.*$",'',ln)
        if ln.strip().startswith('Attribute '): ln=''
        lines.append(ln)
    code='\n'.join(lines).replace(' _\n',' ')
    KW=set('''If Then Else ElseIf End Sub Function Private Public Dim As Long Double String Boolean Variant Range Worksheet Shape Set Nothing Is Not And Or For To Step Next Each In Exit Do Loop While Wend With ByVal ByRef Optional Call New True False Option Explicit Const Mod On Error GoTo Resume Integer Byte Object ReDim Preserve Empty Null Me Select Case
    MsgBox vbExclamation vbInformation vbString Trim Left Mid Len CStr CDbl IsError IsEmpty IsNumeric VarType RGB ChrW Chr UBound LBound IIf Application ThisWorkbook ActiveSheet Err Array Fail Worksheets Abs Int Str Val'''.split())
    gl=set(extra_globals)
    gl|=set(re.findall(r'^(?:Private|Public)\s+(?:Sub|Function|Const)\s+(\w+)',code,re.M))
    fp=re.search(r'^(?:Private |Public )?(?:Sub|Function) ',code,re.M)
    first_proc=fp.start() if fp else len(code)
    for m in re.findall(r'^(?:Private|Public)\s+(?!Sub|Function|Const)(.+)$',code[:first_proc],re.M):
        for part in m.split(','):
            mm=re.match(r'\s*(\w+)',part)
            if mm: gl.add(mm.group(1))
    issues=[('<module>','declaration after first procedure: '+l) for l in re.findall(r'^(?:Private|Public|Dim)\s+(?!Sub|Function|Const).+$',code[first_proc:],re.M)]
    for m in re.finditer(r'^(?:Private |Public )?(?:Sub|Function) (\w+)\(([^\n]*)\)(?: As \w+)?[ \t]*\n(.*?)^End (?:Sub|Function)',code,re.M|re.S):
        name,params,body=m.groups(); loc={name}
        for prm in params.split(','):
            prm=prm.replace('Optional','').replace('ByVal','').replace('ByRef','').strip()
            mm=re.match(r'(\w+)',prm)
            if mm: loc.add(mm.group(1))
        for d in re.findall(r'\bDim\s+([^\n:]+)',body):
            for part in d.split(','):
                mm=re.match(r'\s*(\w+)',part)
                if mm: loc.add(mm.group(1))
        body2=re.sub(r'\.\s*\w+','.X',body); body2=re.sub(r'&H[0-9A-Fa-f]+','0',body2)
        for tok in set(re.findall(r'\b([A-Za-z_]\w*)\$?',body2)):
            if tok in KW or tok in loc or tok in gl or tok=='X': continue
            issues.append((name,tok))
    return issues
if __name__=='__main__':
    print('Module1:', check('module1.bas',['Module1']) or 'OK — every identifier declared, no declarations after procedures')
    print('Sheet1 :', check('sheet1_events.bas',['Module1','Intersect']) or 'OK')
