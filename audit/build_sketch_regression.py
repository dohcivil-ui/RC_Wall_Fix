from pathlib import Path
import re
P=Path(__file__).resolve().parent
s=(P.parent/'RC_RT_HCA_v2.vbp').read_bytes().decode('latin1').replace('\r\n','\n')
s=s.replace('Startup="Form1"','Startup="Sub Main"').replace('Name="Project1"','Name="SketchTest"')
s=re.sub(r'(?m)^(Module=[^;]+; |Form=)([^\n]+)',r'\1..\\\2',s)
s+='\nModule=SketchRegression; SketchRegression.bas\nExeName32="SketchRegression.exe"\n'
(P/'SketchRegression.vbp').write_bytes(s.replace('\n','\r\n').encode('latin1'))
(P/'SketchRegression.bas').write_bytes((P/'SketchRegression.source').read_text(encoding='ascii').replace('\n','\r\n').encode('ascii'))
