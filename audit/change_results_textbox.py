"""Small, byte-preserving migration of the result control to native VB6 TextBox."""
from pathlib import Path
import re

root = Path(__file__).resolve().parent.parent
path = root / 'Form1.frm'
src = path.read_bytes()
assert src.count(b'Begin VB.ListBox lstResults') == 1
src = src.replace(b'Begin VB.ListBox lstResults', b'Begin VB.TextBox txtResults')
src = src.replace(b'Begin VB.TextBox txtResults \r\n',
    b'Begin VB.TextBox txtResults \r\n'
    b'         MultiLine       =   -1  \'True\r\n'
    b'         ScrollBars      =   2  \'Vertical\r\n'
    b'         Locked          =   -1  \'True\r\n'
    b'         Enabled         =   -1  \'True\r\n')
src = src.replace(b'lstResults.Clear', b'txtResults.Text = vbNullString')
src = src.replace(b'lstResults.TopIndex = 0', b'txtResults.SelStart = 0')
src, count = re.subn(rb'Public Sub AddResultLine\(ByVal text As String\)[\s\S]*?End Sub',
    lambda _: b'Public Sub AddResultLine(ByVal text As String)\r\n'
    b'    \' Native multiline TextBox wraps visually; preserve the report text.\r\n'
    b'    txtResults.SelStart = Len(txtResults.Text)\r\n'
    b'    txtResults.SelLength = 0\r\n'
    b'    txtResults.SelText = text & vbCrLf\r\n'
    b'End Sub', src)
assert count == 1 and b'lstResults' not in src
assert src.count(b'\n') == src.count(b'\r\n')
path.write_bytes(src)
