Option Explicit
Function CanonicalDecimal(ByVal s)
    Dim neg
    neg = (Left(s, 1) = "-")
    If neg Then s = Mid(s, 2)
    If InStr(s, ".") > 0 Then
        Do While Right(s, 1) = "0"
            s = Left(s, Len(s) - 1)
        Loop
        If Right(s, 1) = "." Then s = Left(s, Len(s) - 1)
    End If
    Do While Len(s) > 1 And Left(s, 1) = "0" And Mid(s, 2, 1) <> "."
        s = Mid(s, 2)
    Loop
    If neg And s <> "0" Then s = "-" & s
    CanonicalDecimal = s
End Function

Function Matches(ByVal s, ByVal pattern)
    Dim rx
    Set rx = CreateObject("VBScript.RegExp")
    rx.pattern = pattern
    rx.Global = False
    Matches = rx.Test(s)
End Function

Function EscapeXML(ByVal s)
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, """", "&quot;")
    s = Replace(s, vbCr, "&#13;")
    s = Replace(s, vbLf, "&#10;")
    s = Replace(s, vbTab, "&#9;")
    EscapeXML = s
End Function

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
