"""Execute a compatible subset of the actual VBA helpers using Windows VBScript.
Only VBA type declarations and $ string-function suffixes are removed. No Excel launch.
This is NOT a VBA compilation or an Excel end-to-end test.
"""
from pathlib import Path
import re, subprocess
base=Path(__file__).resolve().parents[1]
source=(base/'tools/CBI_REQ_Export_core.bas').read_text()
names=['CanonicalDecimal','Matches','EscapeXML']
functions=[]
for name in names:
    text=re.search(r'Private Function '+name+r'\b.*?End Function',source,re.S).group()
    text=text.replace('Private Function','Function')
    text=re.sub(r' As (String|Boolean|Object|Long|Double|Variant)\b','',text)
    text=re.sub(r'\b(Left|Right|Mid|Trim)\$',r'\1',text)
    functions.append(text)
tests='''
Sub AssertEqual(actual, expected, label)
    If actual <> expected Then
        WScript.Echo "FAIL " & label & " expected=" & expected & " actual=" & actual
        WScript.Quit 1
    End If
End Sub
AssertEqual CanonicalDecimal("000.047000"), "0.047", "percent canonicalization"
AssertEqual CanonicalDecimal("-000.000"), "0", "negative zero"
AssertEqual CanonicalDecimal("123456789012345.123456"), "123456789012345.123456", "large text decimal precision"
AssertEqual CanonicalDecimal("-002.100000"), "-2.1", "negative decimal"
AssertEqual Matches("0.0", "^[0-9]+([.]0+)?$"), True, "integer decimal zeros"
AssertEqual Matches("0.1", "^[0-9]+([.]0+)?$"), False, "integer fractional value"
AssertEqual Matches("1,234.56", "^-?[0-9]+([.,][0-9]+)?$"), False, "ambiguous separators"
AssertEqual Matches("1e3", "^-?[0-9]+([.,][0-9]+)?$"), False, "exponent text rejected"
AssertEqual EscapeXML("A&B <C> " & Chr(34) & "Q" & Chr(34)), "A&amp;B &lt;C&gt; &quot;Q&quot;", "XML escaping"
AssertEqual EscapeXML(vbCrLf & vbTab), "&#13;&#10;&#9;", "attribute whitespace escaping"
WScript.Echo "PASS: 10 source-helper checks in VBScript (not VBA/Excel execution)"
'''
path=base/'tools/qa/helpers.vbs'
path.write_text('Option Explicit\n'+'\n\n'.join(functions)+'\n'+tests,encoding='ascii')
subprocess.run(['C:/Windows/System32/cscript.exe','//nologo',str(path)],check=True)
