Attribute VB_Name = "CBI_REQ_Export"
Option Explicit

' CBI Fund Administrators REQ v1.0. Windows desktop Excel, 32/64 bit.
' Import into the workbook containing the original CBI worksheet layout.
' No external VBA references required. No changes to source answers.
Private Const XMLNS As String = "http://www.centralbank.ie/Schemas/AML/2025/A12"
Private Const XSDNAME As String = "Fund Admin - XSD v1.0.xsd"
Private mSchema As Object
Private mErrors As Collection
Private mOrigin As Collection

Public Sub Eksportuj_XML_CBI()
    RunInteractive True
End Sub

Public Sub Sprawdz_Dane_CBI()
    RunInteractive False
End Sub

Private Sub RunInteractive(ByVal saveXML As Boolean)
    Dim schemaPath As String, reportPath As String, xmlText As String
    Dim code As String, reportDate As String, target As Variant, doc As Object
    Dim stamp As String
    On Error GoTo Failed
    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "Najpierw zapisz skoroszyt jako XLSM.", vbExclamation
        Exit Sub
    End If
    schemaPath = ThisWorkbook.Path & Application.PathSeparator & XSDNAME
    If Len(Dir$(schemaPath)) = 0 Then
        target = Application.GetOpenFilename("Schemat XSD (*.xsd),*.xsd", , "Wybierz oryginalny XSD CBI v1.0")
        If VarType(target) = vbBoolean Then Exit Sub
        schemaPath = CStr(target)
    End If
    Application.StatusBar = "CBI: sprawdzanie danych i tworzenie XML..."
    xmlText = BuildREQ(ThisWorkbook, schemaPath)
    Application.StatusBar = False
    stamp = Format$(Now, "yyyymmdd_hhnnss")
    reportPath = ThisWorkbook.Path & Application.PathSeparator & "CBI_validation_" & stamp & ".txt"
    If mErrors.Count > 0 Then
        WriteUTF8 reportPath, ErrorReport()
        MsgBox "Nie zapisano XML. Liczba bledow: " & mErrors.Count & vbCrLf & _
               "Raport: " & reportPath & vbCrLf & vbCrLf & CStr(mErrors(1)), vbExclamation
        Exit Sub
    End If
    If Not saveXML Then
        MsgBox "Kontrola kompletnosci technicznej i walidacja XSD zakonczone poprawnie." & vbCrLf & _
               "Sprawdz jeszcze poprawnosc merytoryczna danych i usun dane przykladowe banku.", vbInformation
        Exit Sub
    End If
    code = Trim$(InputBox("Kod instytucji uzywany w Portalu CBI (dokladnie taki jak w Portalu):", "Nazwa pliku XML"))
    If Len(code) = 0 Then Exit Sub
    If Not Matches(code, "^[A-Za-z0-9]+$") Then
        MsgBox "Kod instytucji moze zawierac tylko litery i cyfry.", vbExclamation
        Exit Sub
    End If
    reportDate = Trim$(InputBox("Data raportowa YYYY-MM-DD, np. 2026-12-31. Nie data wysylki:", "Data raportowa"))
    If Len(reportDate) = 0 Then Exit Sub
    If Not ValidISODate(reportDate) Then
        MsgBox "Nieprawidlowa data raportowa (wymagany zakres 2000-2099).", vbExclamation
        Exit Sub
    End If
    target = Application.GetSaveAsFilename(ThisWorkbook.Path & Application.PathSeparator & _
             code & "_" & Replace(reportDate, "-", "") & "_A12.xml", "XML (*.xml),*.xml")
    If VarType(target) = vbBoolean Then Exit Sub
    If LCase$(CreateObject("Scripting.FileSystemObject").GetFileName(CStr(target))) <> _
       LCase$(code & "_" & Replace(reportDate, "-", "") & "_A12.xml") Then
        MsgBox "Zachowaj wygenerowana nazwe: kod_YYYYMMDD_A12.xml. Uruchom eksport ponownie.", vbExclamation
        Exit Sub
    End If
    If Len(Dir$(CStr(target))) > 0 Then
        If MsgBox("Plik juz istnieje. Zastapic go?" & vbCrLf & CStr(target), vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    End If
    WriteUTF8 CStr(target), xmlText
    MsgBox "Zapisano XML zgodny z XSD:" & vbCrLf & CStr(target) & vbCrLf & _
           "Po sprawdzeniu danych przeslij plik przez Portal CBI oraz wykonaj Finalise i Sign-Off.", vbInformation
    Exit Sub
Failed:
    Application.StatusBar = False
    MsgBox "Eksport przerwany: " & Err.Description, vbCritical
End Sub

' Public entry point for controlled automated tests; never saves a workbook.
' Returns an empty string on validation failure; errors available via CBI_LastErrors.
Public Function BuildREQ(ByVal wb As Object, ByVal schemaPath As String) As String
    Dim scalarMap As New Collection, tableMap As New Collection
    Dim schemaCache As Object, sequence As Object, groupSpec As Object
    Dim groupName As String, item As Variant, p As Variant, body As String, groupBody As String
    Dim dom As Object, validation As Object, check As Object, origin As String, sh As Object
    Set mErrors = New Collection
    Set mOrigin = New Collection
    On Error GoTo Failed
    Set mSchema = NewDOM()
    If Not mSchema.Load(schemaPath) Then Err.Raise vbObjectError + 100, , "Nie mozna odczytac XSD: " & mSchema.parseError.reason
    mSchema.setProperty "SelectionNamespaces", "xmlns:xs='http://www.w3.org/2001/XMLSchema'"
    If mSchema.documentElement.getAttribute("targetNamespace") <> XMLNS Then Err.Raise vbObjectError + 101, , "Niewlasciwa przestrzen nazw XSD. Wymagany FundAdmin A12 v1.0."
    If mSchema.documentElement.getAttribute("version") <> "1.0" Then Err.Raise vbObjectError + 102, , "Eksporter obsluguje XSD v1.0."
    Set schemaCache = CreateObject("MSXML2.XMLSchemaCache.6.0")
    schemaCache.Add XMLNS, mSchema
    LoadScalarMap scalarMap
    LoadTableMap tableMap
    If Not CheckLayout(wb, scalarMap, tableMap) Then Exit Function
    For Each sh In wb.Worksheets
        sh.Calculate
    Next sh
    body = "<?xml version=""1.0"" encoding=""utf-8""?>" & vbCrLf
    body = body & "<FundAdmin xmlns=""" & XMLNS & """>" & vbCrLf
    Set sequence = mSchema.selectNodes("/xs:schema/xs:element[@name='FundAdmin']/xs:complexType/xs:sequence/xs:element")
    If sequence.Length <> 12 Then Err.Raise vbObjectError + 103, , "Nieoczekiwana struktura XSD."
    For Each groupSpec In sequence
        groupName = Split(groupSpec.getAttribute("ref"), ":")(1)
        body = body & "  <" & groupName & ">" & vbCrLf
        groupBody = ""
        For Each item In scalarMap
            p = Split(CStr(item), "|")
            If p(3) = groupName Then groupBody = groupBody & ScalarXML(wb, p)
        Next item
        For Each item In tableMap
            p = Split(CStr(item), "|")
            If p(0) = groupName Then groupBody = groupBody & TableXML(wb, p)
        Next item
        body = body & groupBody & "  </" & groupName & ">" & vbCrLf
    Next groupSpec
    body = body & "</FundAdmin>" & vbCrLf
    If mErrors.Count > 0 Then Exit Function
    Set dom = NewDOM()
    Set dom.schemas = schemaCache
    If Not dom.LoadXML(body) Then
        AddError "XML", dom.parseError.reason
        Exit Function
    End If
    Set validation = dom.validate
    If validation.errorCode <> 0 Then
        origin = "XML / XSD"
        If validation.line > 0 Then origin = origin & " / linia " & validation.line
        AddError origin, validation.reason
        ' Locate XSD failures down to the source row without changing the XML.
        Dim nodes As Object, n As Object, pos As Long
        Set nodes = dom.selectNodes("/*/*/*")
        pos = 0
        For Each n In nodes
            pos = pos + 1
            Set check = dom.validateNode(n)
            If check.errorCode <> 0 Then AddError CStr(mOrigin(pos)), check.reason
        Next n
        Exit Function
    End If
    BuildREQ = body
    Exit Function
Failed:
    AddError "Eksporter", Err.Description
End Function

Public Function CBI_LastErrors() As String
    CBI_LastErrors = ErrorReport()
End Function

Private Function NewDOM() As Object
    Dim d As Object
    Set d = CreateObject("MSXML2.DOMDocument.6.0")
    d.async = False
    d.validateOnParse = False
    d.resolveExternals = False
    d.setProperty "ProhibitDTD", True
    Set NewDOM = d
End Function

Private Function CheckLayout(ByVal wb As Object, ByVal scalars As Collection, ByVal tables As Collection) As Boolean
    Dim item As Variant, p As Variant, sh As Object, cols As Variant, refs As Variant, i As Long
    On Error GoTo BadLayout
    For Each item In scalars
        p = Split(CStr(item), "|")
        Set sh = wb.Worksheets(p(1))
        If LabelHash(CStr(sh.Range(p(6)).Value2)) <> p(7) Then
            AddError p(1) & "!" & p(6), "Zmieniony uklad lub tresc etykiety. Uzyj oryginalnego ukladu CBI v1.0."
        End If
    Next item
    For Each item In tables
        p = Split(CStr(item), "|")
        Set sh = wb.Worksheets(p(1))
        cols = Split(p(5), ",")
        refs = Split(p(6), ",")
        For i = 0 To UBound(cols)
            If CStr(sh.Range(cols(i) & p(2)).Value2) <> refs(i) Then
                AddError p(1) & "!" & cols(i) & p(2), "Brak naglowka " & refs(i) & ". Nie przesuwaj kolumn ani wierszy naglowka."
            End If
        Next i
    Next item
    CheckLayout = (mErrors.Count = 0)
    Exit Function
BadLayout:
    AddError "Uklad skoroszytu", "Brak arkusza lub niezgodny szablon: " & Err.Description
End Function

Private Function LabelHash(ByVal s As String) As String
    Dim i As Long, h As Double, u As Long
    For i = 1 To Len(s)
        u = AscW(Mid$(s, i, 1))
        If u < 0 Then u = u + 65536
        h = h * 31# + u
        h = h - Fix(h / 2147483647#) * 2147483647#
    Next i
    LabelHash = Format$(h, "0")
End Function

Private Function ScalarXML(ByVal wb As Object, ByVal p As Variant) As String
    Dim cell As Object, val As String, typ As String, place As String
    Set cell = wb.Worksheets(p(1)).Range(p(2))
    place = p(0) & " / " & p(1) & "!" & p(2)
    typ = AttributeType(CStr(p(3)), CStr(p(5)))
    val = CellValue(cell, typ, CStr(p(0)), place, wb.Date1904)
    ScalarXML = "    <" & p(4) & " Identifier=""" & p(0) & """ " & p(5) & "=""" & EscapeXML(val) & """></" & p(4) & ">" & vbCrLf
    mOrigin.Add place
End Function

Private Function TableXML(ByVal wb As Object, ByVal p As Variant) As String
    Dim sh As Object, cols As Variant, refs As Variant, spec As Object
    Dim firstRow As Long, lastRow As Long, r As Long, i As Long, count As Long
    Dim key As String, val As String, attrs As String, typ As String, place As String
    Dim element As String, result As String, used As Boolean, seen As Object, maxRows As String
    Set sh = wb.Worksheets(p(1))
    cols = Split(p(5), ",")
    refs = Split(p(6), ",")
    firstRow = CLng(p(3))
    lastRow = CLng(p(4))
    If lastRow = 0 Then lastRow = LastDataRow(sh, CStr(cols(0)), CStr(cols(UBound(cols))), firstRow)
    Set spec = mSchema.selectSingleNode("/xs:schema/xs:element[@name='" & p(0) & "']/xs:complexType/xs:sequence/xs:element")
    element = spec.getAttribute("name")
    Set seen = CreateObject("Scripting.Dictionary")
    For r = firstRow To lastRow
        used = False
        For i = 0 To UBound(cols)
            If IsError(sh.Range(cols(i) & r).Value2) Then
                used = True
            ElseIf Len(Trim$(CStr(sh.Range(cols(i) & r).Value2))) > 0 Then
                used = True
            End If
        Next i
        If used Or CLng(p(4)) <> 0 Then
            count = count + 1
            attrs = ""
            For i = 0 To UBound(cols)
                place = refs(i) & " / " & p(1) & "!" & cols(i) & r
                typ = AttributeType(CStr(p(0)), CStr(refs(i)))
                val = CellValue(sh.Range(cols(i) & r), typ, CStr(refs(i)), place, wb.Date1904)
                If i = 0 Then key = val
                attrs = attrs & " " & refs(i) & "=""" & EscapeXML(val) & """"
            Next i
            If Len(key) > 0 Then
                If seen.Exists(key) Then
                    AddError p(1) & "!" & cols(0) & r, "Powtorzony klucz wiersza: " & key & "; poprzedni wiersz: " & seen(key)
                Else
                    seen.Add key, r
                End If
            End If
            result = result & "    <" & element & attrs & "></" & element & ">" & vbCrLf
            mOrigin.Add p(1) & " / wiersz " & r & " / klucz " & key
        End If
    Next r
    If count < CLng(spec.getAttribute("minOccurs")) Then AddError CStr(p(1)), "Za malo wierszy. Wypelnij tabele; jesli nie dotyczy, wpisz jawnie wartosci zgodnie z instrukcja CBI."
    maxRows = spec.getAttribute("maxOccurs")
    If maxRows <> "unbounded" Then
        If count > CLng(maxRows) Then AddError CStr(p(1)), "Za duzo wierszy: " & count & ". Limit XSD: " & maxRows
    End If
    If p(0) <> "A1202A" And p(0) <> "A1209" Then
        If seen.Exists("00") And count > 1 Then AddError CStr(p(1)), "Kod 00 (nie dotyczy) nie moze wystepowac razem z innymi krajami."
    End If
    TableXML = result
End Function

Private Function LastDataRow(ByVal sh As Object, ByVal firstCol As String, ByVal lastCol As String, ByVal startRow As Long) As Long
    Dim region As Object, found As Object
    Set region = sh.Range(firstCol & startRow & ":" & lastCol & sh.Rows.Count)
    Set found = region.Find(What:="*", After:=region.Cells(1, 1), LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
    LastDataRow = startRow - 1
    If Not found Is Nothing Then LastDataRow = found.Row
End Function

Private Function AttributeType(ByVal group As String, ByVal attr As String) As String
    Dim n As Object
    Set n = mSchema.selectSingleNode("/xs:schema/xs:element[@name='" & group & "']//xs:attribute[@name='" & attr & "']")
    If n Is Nothing Then Err.Raise vbObjectError + 110, , "XSD nie zawiera " & group & "/@" & attr
    AttributeType = Split(n.getAttribute("type"), ":")(1)
End Function

Private Function CellValue(ByVal cell As Object, ByVal typ As String, ByVal ref As String, ByVal place As String, ByVal date1904 As Boolean) As String
    Dim v As Variant, s As String, n As Double, restriction As Object, opts As Object, opt As Object
    Dim allowed As String, good As Boolean, limit As Long, dt As Date
    v = cell.Value2
    If IsError(v) Then
        AddError place, "Blad formuly w Excelu."
        Exit Function
    End If
    s = Trim$(CStr(v))
    If Len(s) = 0 Or s = ref Or Left$(s, 1) = "<" Then
        AddError place, "Brak odpowiedzi lub pozostawiony kod/podpowiedz banku."
        Exit Function
    End If
    If VarType(v) = vbBoolean Then
        AddError place, "Wartosc logiczna Excel nie jest odpowiedzia CBI. Wpisz wymagany kod lub liczbe."
        Exit Function
    End If
    Set restriction = mSchema.selectSingleNode("/xs:schema/xs:simpleType[@name='" & typ & "']/xs:restriction")
    If restriction Is Nothing Then Err.Raise vbObjectError + 111, , "Nieznany typ XSD: " & typ
    Set opts = restriction.selectNodes("xs:enumeration")
    If opts.Length > 0 Then
        For Each opt In opts
            If s = opt.getAttribute("value") Then good = True
            If Len(allowed) < 240 Then allowed = allowed & opt.getAttribute("value") & ", "
        Next opt
        If Not good Then AddError place, "Niedozwolona odpowiedz '" & s & "' dla " & typ & ". Dozwolone (poczatek listy): " & allowed
    Else
        Select Case typ
        Case "NoNegInt20", "Decimal206", "Percentage"
            If VarType(v) = vbString Then
                ' Text numbers: no grouping or exponent; accept either decimal mark.
                If Not Matches(s, "^-?[0-9]+([.,][0-9]+)?$") Then
                    AddError place, "Nieprawidlowa liczba. Bez separatorow tysiecy; procent podaj jako ulamek (0.047)."
                    Exit Function
                End If
                s = Replace(s, ",", ".")
            ElseIf IsNumeric(v) Then
                If Abs(CDbl(v)) >= 1000000000000000# Then
                    AddError place, "Liczba ma co najmniej 16 cyfr przed kropka. Wpisz ja jako tekst, aby uniknac utraty precyzji Excela."
                    Exit Function
                End If
                If CDbl(v) <> 0 And Abs(CDbl(v)) < 0.000001 Then
                    AddError place, "Wartosc wymaga ponad 6 miejsc po przecinku. Nie zaokraglam danych automatycznie."
                    Exit Function
                End If
                ' CDec avoids exponent notation without fixed-format rounding.
                s = Replace(CStr(CDec(v)), Application.International(xlDecimalSeparator), ".")
            Else
                AddError place, "Oczekiwana liczba."
                Exit Function
            End If
            If typ = "NoNegInt20" Then
                n = Val(s)
                If n < 0 Or n > 2147483647# Or Not Matches(s, "^[0-9]+([.]0+)?$") Then
                    AddError place, "Wymagana liczba calkowita 0..2147483647."
                Else
                    s = Format$(n, "0")
                End If
            Else
                s = CanonicalDecimal(s)
                If InStr(s, ".") > 0 Then
                    If Len(s) - InStr(s, ".") > 6 Then AddError place, "Maksymalnie 6 miejsc po przecinku. Nie zaokraglam danych automatycznie."
                End If
                If typ = "Percentage" Then
                    If Val(s) < 0 Or Val(s) > 1 Then AddError place, "Procent musi byc ulamkiem od 0 do 1 (4.7% = 0.047)."
                End If
            End If
        Case "ReportDate"
            If VarType(v) <> vbString And IsNumeric(v) Then
                n = CDbl(v)
                If n <> Fix(n) Or n < 0 Or n > 73050 Then
                    AddError place, "Nieprawidlowa data Excela lub data zawiera godzine."
                    Exit Function
                End If
                If date1904 Then
                    dt = DateSerial(1904, 1, 1) + n
                Else
                    dt = DateSerial(1899, 12, 30) + n
                End If
                s = Format$(dt, "yyyy-mm-dd")
            End If
            If Not ValidISODate(s) Then AddError place, "Wymagana data YYYY-MM-DD od 2000-01-01 do 2099-12-31."
        Case "LEI"
            If VarType(v) <> vbString Then
                AddError place, "LEI wpisz jako tekst o 20 znakach (takze 20 zer dla nie dotyczy)."
            ElseIf Not Matches(s, "^[A-Z0-9]{20}$") Then
                AddError place, "LEI musi zawierac 20 duzych liter/cyfr."
            End If
        Case "String100", "String500", "String2000"
            ' Guidance limits are stricter than the XSD's +50 character allowance.
            limit = CLng(Mid$(typ, 7))
            If Len(s) > limit Then AddError place, "Przekroczony limit instrukcji: " & limit & " znakow."
        Case Else
            Err.Raise vbObjectError + 112, , "Nieobslugiwany typ: " & typ
        End Select
    End If
    CellValue = s
End Function

Private Function CanonicalDecimal(ByVal s As String) As String
    Dim neg As Boolean
    neg = (Left$(s, 1) = "-")
    If neg Then s = Mid$(s, 2)
    If InStr(s, ".") > 0 Then
        Do While Right$(s, 1) = "0"
            s = Left$(s, Len(s) - 1)
        Loop
        If Right$(s, 1) = "." Then s = Left$(s, Len(s) - 1)
    End If
    Do While Len(s) > 1 And Left$(s, 1) = "0" And Mid$(s, 2, 1) <> "."
        s = Mid$(s, 2)
    Loop
    If neg And s <> "0" Then s = "-" & s
    CanonicalDecimal = s
End Function

Private Function ValidISODate(ByVal s As String) As Boolean
    Dim d As Date
    On Error GoTo InvalidDate
    If Not Matches(s, "^[0-9]{4}-[0-9]{2}-[0-9]{2}$") Then Exit Function
    If Left$(s, 4) < "2000" Or Left$(s, 4) > "2099" Then Exit Function
    d = DateSerial(CInt(Left$(s, 4)), CInt(Mid$(s, 6, 2)), CInt(Right$(s, 2)))
    ValidISODate = (Format$(d, "yyyy-mm-dd") = s)
InvalidDate:
End Function

Private Function Matches(ByVal s As String, ByVal pattern As String) As Boolean
    Dim rx As Object
    Set rx = CreateObject("VBScript.RegExp")
    rx.pattern = pattern
    rx.Global = False
    Matches = rx.Test(s)
End Function

Private Function EscapeXML(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, """", "&quot;")
    s = Replace(s, vbCr, "&#13;")
    s = Replace(s, vbLf, "&#10;")
    s = Replace(s, vbTab, "&#9;")
    EscapeXML = s
End Function

Private Sub AddError(ByVal place As String, ByVal reason As String)
    mErrors.Add place & ": " & Trim$(reason)
End Sub

Private Function ErrorReport() As String
    Dim item As Variant, s As String
    s = "CBI Fund Administrators v1.0 - raport walidacji" & vbCrLf & _
        "Data: " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & _
        "Nie zapisano nowego XML. Ewentualny plik z poprzedniego eksportu pozostaje niezmieniony." & vbCrLf & vbCrLf
    If mErrors Is Nothing Then
        ErrorReport = "Nie uruchomiono walidacji."
        Exit Function
    End If
    For Each item In mErrors
        s = s & CStr(item) & vbCrLf
    Next item
    ErrorReport = s
End Function

Private Sub WriteUTF8(ByVal path As String, ByVal text As String)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText text
    stream.SaveToFile path, 2
    stream.Close
End Sub
