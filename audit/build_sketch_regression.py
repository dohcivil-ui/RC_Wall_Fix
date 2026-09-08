from pathlib import Path
import re, sys
P=Path(__file__).resolve().parent
s=(P.parent/'RC_RT_HCA_v2.vbp').read_bytes().decode('latin1').replace('\r\n','\n')
s=s.replace('Startup="Form1"','Startup="Sub Main"').replace('Name="Project1"','Name="SketchTest"')
s=re.sub(r'(?m)^(Module=[^;]+; |Form=)([^\n]+)',r'\1..\\\2',s)
s+='\nModule=SketchRegression; SketchRegression.bas\nModule=WindowCapture; WindowCapture.bas\nExeName32="SketchRegression.exe"\n'
(P/'SketchRegression.vbp').write_bytes(s.replace('\n','\r\n').encode('latin1'))
test=(P/'SketchRegression.source').read_text(encoding='ascii')
if '--single-trial' in sys.argv:
    test=test.replace('Form1.txtTrials.Text = "2"','Form1.txtTrials.Text = "1"')
    test=test.replace('RunTrial = 2 And EvaluationCount = 64','RunTrial = 1 And EvaluationCount = 64')
    test=test.replace('completed both trials','completed single trial')
    test=test.replace('    CheckSession "HCA"\n','')
(P/'SketchRegression.bas').write_bytes(test.replace('\n','\r\n').encode('ascii'))
