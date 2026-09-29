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

Private Sub LoadScalarMap(ByRef records As Collection)
    MapPart0 records
    MapPart1 records
    MapPart2 records
    MapPart3 records
    MapPart4 records
    MapPart5 records
    MapPart6 records
    MapPart7 records
End Sub

Private Sub MapPart0(ByRef records As Collection)
    records.Add "LAA|8.2 General|D3|A1201A|A12011|EnumerationLegalStructure|B3|840638667"
    records.Add "LAB|8.2 General|D5|A1201A|A12011|String500|B5|1774445322"
    records.Add "LAC|8.2 General|D6|A1201A|A12011|LEI|B6|892179086"
    records.Add "LAD|8.2 General|D7|A1201A|A12011|EnumerationCountryEEANA|B7|1692520932"
    records.Add "LAE|8.2 General|D8|A1201A|A12011|String500|B8|1131924430"
    records.Add "LAF|8.2 General|D9|A1201A|A12011|String500|B9|1261450077"
    records.Add "LAG|8.2 General|D10|A1201A|A12011|LEI|B10|720369034"
    records.Add "LAH|8.2 General|D11|A1201A|A12011|EnumerationCountry|B11|855172215"
    records.Add "LAI|8.2 General|D12|A1201A|A12011|String500|B12|443302216"
    records.Add "LAJ|8.2 General|D14|A1201A|A12011|String500|B14|1549903747"
    records.Add "LAK|8.2 General|D15|A1201A|A12011|LEI|B15|718613737"
    records.Add "LAL|8.2 General|D16|A1201A|A12011|EnumerationCountryEEANA|B16|282810737"
    records.Add "LAM|8.2 General|D17|A1201A|A12011|String500|B17|206302917"
    records.Add "LAN|8.2 General|D18|A1201A|A12011|String500|B18|1591119992"
    records.Add "LAO|8.2 General|D19|A1201A|A12011|LEI|B19|1411849286"
    records.Add "LAP|8.2 General|D20|A1201A|A12011|EnumerationCountry|B20|1670635785"
    records.Add "LAQ|8.2 General|D21|A1201A|A12011|String500|B21|1326489895"
    records.Add "LAW|8.2 General|D62|A1201A|A12011|EnumerationYesNo|B62|887531050"
    records.Add "LAX|8.2 General|D65|A1201A|A12011|String500|B65|312600406"
    records.Add "LAY|8.2 General|D68|A1201A|A12011|String100|B68|1692000966"
    records.Add "LAZ|8.2 General|E68|A1201A|A12011|String100|B68|1692000966"
    records.Add "LBA|8.2 General|F68|A1201A|A12011|EnumerationYesNo|B68|1692000966"
    records.Add "LBS|8.3.2 Inherent Risks Investors|D40|A1202C|A12023|EnumerationYesNo|B40|2073630294"
    records.Add "LBT|8.3.2 Inherent Risks Investors|D42|A1202C|A12023|IntegerValue|B42|1280927121"
    records.Add "LBU|8.3.2 Inherent Risks Investors|D43|A1202C|A12023|IntegerValue|B43|1682639584"
    records.Add "LBV|8.4 Mitigation & Control|C5|A1203A|A12031|EnumerationBWRAConducted|B5|738847092"
    records.Add "LBW|8.4 Mitigation & Control|C8|A1203A|A12031|EnumerationYesNo|B8|739770534"
    records.Add "LBX|8.4 Mitigation & Control|C9|A1203A|A12031|DateValue|B9|1210042584"
    records.Add "LBY|8.4 Mitigation & Control|C10|A1203A|A12031|EnumerationFrequency|B10|1471149071"
    records.Add "LBZ|8.4 Mitigation & Control|C11|A1203A|A12031|EnumerationYesNoNA|B11|1553913321"
    records.Add "LCA|8.4 Mitigation & Control|C12|A1203A|A12031|String500|B12|831384861"
    records.Add "LCB|8.4 Mitigation & Control|C16|A1203A|A12031|EnumerationYesNoNA|B16|1899793439"
    records.Add "LCC|8.4 Mitigation & Control|C17|A1203A|A12031|EnumerationYesNoNA|B17|1837590984"
    records.Add "LCD|8.4 Mitigation & Control|C18|A1203A|A12031|EnumerationYesNoNA|B18|1025056210"
    records.Add "LCE|8.4 Mitigation & Control|C19|A1203A|A12031|EnumerationYesNoNA|B19|845140119"
    records.Add "LCF|8.4 Mitigation & Control|C20|A1203A|A12031|EnumerationYesNoNA|B20|1652532959"
    records.Add "LCG|8.4 Mitigation & Control|C21|A1203A|A12031|EnumerationYesNoNA|B21|765175071"
    records.Add "LCH|8.4 Mitigation & Control|C22|A1203A|A12031|EnumerationYesNoNA|B22|1694100254"
    records.Add "LCI|8.4 Mitigation & Control|C23|A1203A|A12031|EnumerationYesNoNA|B23|1705557943"
    records.Add "LCJ|8.4 Mitigation & Control|C24|A1203A|A12031|EnumerationYesNoNA|B24|1887047733"
    records.Add "LCK|8.4 Mitigation & Control|C25|A1203A|A12031|EnumerationYesNoNA|B25|1699186303"
    records.Add "LCL|8.4 Mitigation & Control|C26|A1203A|A12031|EnumerationYesNoNA|B26|1948497788"
    records.Add "LCM|8.4 Mitigation & Control|C27|A1203A|A12031|EnumerationYesNoNA|B27|2110073219"
    records.Add "LCN|8.4 Mitigation & Control|C28|A1203A|A12031|EnumerationYesNoNA|B28|47977870"
    records.Add "LCO|8.4 Mitigation & Control|C29|A1203A|A12031|EnumerationYesNoNA|B29|495928992"
    records.Add "LCP|8.4 Mitigation & Control|C30|A1203A|A12031|EnumerationYesNoNA|B30|2027150561"
    records.Add "LCQ|8.4 Mitigation & Control|C31|A1203A|A12031|EnumerationYesNoNA|B31|76517104"
    records.Add "LCR|8.4 Mitigation & Control|C32|A1203A|A12031|String500|B32|1493994058"
    records.Add "LCS|8.4 Mitigation & Control|D16|A1203A|A12031|EnumerationYesNoNA|B16|1899793439"
    records.Add "LCT|8.4 Mitigation & Control|D17|A1203A|A12031|EnumerationYesNoNA|B17|1837590984"
    records.Add "LCU|8.4 Mitigation & Control|D18|A1203A|A12031|EnumerationYesNoNA|B18|1025056210"
    records.Add "LCV|8.4 Mitigation & Control|D19|A1203A|A12031|EnumerationYesNoNA|B19|845140119"
    records.Add "LCW|8.4 Mitigation & Control|D20|A1203A|A12031|EnumerationYesNoNA|B20|1652532959"
    records.Add "LCX|8.4 Mitigation & Control|D21|A1203A|A12031|EnumerationYesNoNA|B21|765175071"
    records.Add "LCY|8.4 Mitigation & Control|D22|A1203A|A12031|EnumerationYesNoNA|B22|1694100254"
    records.Add "LCZ|8.4 Mitigation & Control|D23|A1203A|A12031|EnumerationYesNoNA|B23|1705557943"
    records.Add "LDA|8.4 Mitigation & Control|D24|A1203A|A12031|EnumerationYesNoNA|B24|1887047733"
    records.Add "LDB|8.4 Mitigation & Control|D25|A1203A|A12031|EnumerationYesNoNA|B25|1699186303"
    records.Add "LDC|8.4 Mitigation & Control|D26|A1203A|A12031|EnumerationYesNoNA|B26|1948497788"
    records.Add "LDD|8.4 Mitigation & Control|D27|A1203A|A12031|EnumerationYesNoNA|B27|2110073219"
End Sub

Private Sub MapPart1(ByRef records As Collection)
    records.Add "LDE|8.4 Mitigation & Control|D28|A1203A|A12031|EnumerationYesNoNA|B28|47977870"
    records.Add "LDF|8.4 Mitigation & Control|D29|A1203A|A12031|EnumerationYesNoNA|B29|495928992"
    records.Add "LDG|8.4 Mitigation & Control|D30|A1203A|A12031|EnumerationYesNoNA|B30|2027150561"
    records.Add "LDH|8.4 Mitigation & Control|D31|A1203A|A12031|EnumerationYesNoNA|B31|76517104"
    records.Add "LDI|8.4 Mitigation & Control|D32|A1203A|A12031|String500|B32|1493994058"
    records.Add "LDJ|8.4 Mitigation & Control|C35|A1203A|A12031|EnumerationYesNoNA|B35|385922308"
    records.Add "LDK|8.4 Mitigation & Control|C36|A1203A|A12031|EnumerationYesNoNA|B36|1892107278"
    records.Add "LDL|8.4 Mitigation & Control|C37|A1203A|A12031|EnumerationYesNoNA|B37|1270555607"
    records.Add "LDM|8.4 Mitigation & Control|C38|A1203A|A12031|EnumerationYesNoNA|B38|277284675"
    records.Add "LDN|8.4 Mitigation & Control|C39|A1203A|A12031|EnumerationYesNoNA|B39|76517104"
    records.Add "LDO|8.4 Mitigation & Control|C40|A1203A|A12031|String500|B40|636548474"
    records.Add "LDP|8.4 Mitigation & Control|C41|A1203A|A12031|EnumerationYesNo|B41|46301148"
    records.Add "LDQ|8.4 Mitigation & Control|C47|A1203A|A12031|DateValue|B47|2015598937"
    records.Add "LDR|8.4 Mitigation & Control|C48|A1203A|A12031|DateValue|B48|4319014"
    records.Add "LDS|8.4 Mitigation & Control|C49|A1203A|A12031|DateValue|B49|1023201054"
    records.Add "LDT|8.4 Mitigation & Control|C50|A1203A|A12031|DateValue|B50|2051726412"
    records.Add "LDU|8.4 Mitigation & Control|C51|A1203A|A12031|DateValue|B51|343610481"
    records.Add "LDV|8.4 Mitigation & Control|C52|A1203A|A12031|DateValue|B52|880441845"
    records.Add "LDW|8.4 Mitigation & Control|C53|A1203A|A12031|DateValue|B53|2007878155"
    records.Add "LDX|8.4 Mitigation & Control|C54|A1203A|A12031|DateValue|B54|509598728"
    records.Add "LDY|8.4 Mitigation & Control|C55|A1203A|A12031|DateValue|B55|1809413983"
    records.Add "LDZ|8.4 Mitigation & Control|C56|A1203A|A12031|DateValue|B56|1679930019"
    records.Add "LEA|8.4 Mitigation & Control|C57|A1203A|A12031|DateValue|B57|180708802"
    records.Add "LEB|8.4 Mitigation & Control|C58|A1203A|A12031|DateValue|B58|1287919970"
    records.Add "LEC|8.4 Mitigation & Control|C59|A1203A|A12031|DateValue|B59|983771579"
    records.Add "LED|8.4 Mitigation & Control|C60|A1203A|A12031|DateValue|B60|1046447086"
    records.Add "LEE|8.4 Mitigation & Control|C61|A1203A|A12031|DateValue|B61|7452496"
    records.Add "LEF|8.4 Mitigation & Control|C62|A1203A|A12031|DateValue|B62|1340764510"
    records.Add "LEG|8.4 Mitigation & Control|C63|A1203A|A12031|DateValue|B63|83251771"
    records.Add "LEH|8.4 Mitigation & Control|C64|A1203A|A12031|DateValue|B64|1146299372"
    records.Add "LEI|8.4 Mitigation & Control|C65|A1203A|A12031|DateValue|B65|819994185"
    records.Add "LEJ|8.4 Mitigation & Control|C66|A1203A|A12031|DateValue|B66|1878767588"
    records.Add "LEK|8.4 Mitigation & Control|D47|A1203A|A12031|DateValue|B47|2015598937"
    records.Add "LEL|8.4 Mitigation & Control|D48|A1203A|A12031|DateValue|B48|4319014"
    records.Add "LEM|8.4 Mitigation & Control|D49|A1203A|A12031|DateValue|B49|1023201054"
    records.Add "LEN|8.4 Mitigation & Control|D50|A1203A|A12031|DateValue|B50|2051726412"
    records.Add "LEO|8.4 Mitigation & Control|D51|A1203A|A12031|DateValue|B51|343610481"
    records.Add "LEP|8.4 Mitigation & Control|D52|A1203A|A12031|DateValue|B52|880441845"
    records.Add "LEQ|8.4 Mitigation & Control|D53|A1203A|A12031|DateValue|B53|2007878155"
    records.Add "LER|8.4 Mitigation & Control|D54|A1203A|A12031|DateValue|B54|509598728"
    records.Add "LES|8.4 Mitigation & Control|D55|A1203A|A12031|DateValue|B55|1809413983"
    records.Add "LET|8.4 Mitigation & Control|D56|A1203A|A12031|DateValue|B56|1679930019"
    records.Add "LEU|8.4 Mitigation & Control|D57|A1203A|A12031|DateValue|B57|180708802"
    records.Add "LEV|8.4 Mitigation & Control|D58|A1203A|A12031|DateValue|B58|1287919970"
    records.Add "LEW|8.4 Mitigation & Control|D59|A1203A|A12031|DateValue|B59|983771579"
    records.Add "LEX|8.4 Mitigation & Control|D60|A1203A|A12031|DateValue|B60|1046447086"
    records.Add "LEY|8.4 Mitigation & Control|D61|A1203A|A12031|DateValue|B61|7452496"
    records.Add "LEZ|8.4 Mitigation & Control|D62|A1203A|A12031|DateValue|B62|1340764510"
    records.Add "LFA|8.4 Mitigation & Control|D63|A1203A|A12031|DateValue|B63|83251771"
    records.Add "LFB|8.4 Mitigation & Control|D64|A1203A|A12031|DateValue|B64|1146299372"
    records.Add "LFC|8.4 Mitigation & Control|D65|A1203A|A12031|DateValue|B65|819994185"
    records.Add "LFD|8.4 Mitigation & Control|D66|A1203A|A12031|DateValue|B66|1878767588"
    records.Add "LFE|8.4 Mitigation & Control|C69|A1203A|A12031|DateValue|B69|1799445586"
    records.Add "LFF|8.4 Mitigation & Control|C70|A1203A|A12031|DateValue|B70|1521768552"
    records.Add "LFG|8.4 Mitigation & Control|C71|A1203A|A12031|DateValue|B71|1904801907"
    records.Add "LFH|8.4 Mitigation & Control|C72|A1203A|A12031|DateValue|B72|2093291191"
    records.Add "LFI|8.4 Mitigation & Control|C73|A1203A|A12031|DateValue|B73|2009781336"
    records.Add "LFJ|8.4 Mitigation & Control|C75|A1203A|A12031|EnumerationFrequency|B75|1083461588"
    records.Add "LFK|8.4 Mitigation & Control|C79|A1203A|A12031|IntegerValue|B79|449572829"
    records.Add "LFL|8.4 Mitigation & Control|C80|A1203A|A12031|IntegerValue|B80|1464100973"
End Sub

Private Sub MapPart2(ByRef records As Collection)
    records.Add "LFM|8.4 Mitigation & Control|C81|A1203A|A12031|IntegerValue|B81|1988283607"
    records.Add "LFN|8.4 Mitigation & Control|C82|A1203A|A12031|IntegerValue|B82|1708997244"
    records.Add "LFO|8.4 Mitigation & Control|C83|A1203A|A12031|IntegerValue|B83|380525302"
    records.Add "LFP|8.4 Mitigation & Control|D79|A1203A|A12031|IntegerValue|B79|449572829"
    records.Add "LFQ|8.4 Mitigation & Control|D80|A1203A|A12031|IntegerValue|B80|1464100973"
    records.Add "LFR|8.4 Mitigation & Control|D81|A1203A|A12031|IntegerValue|B81|1988283607"
    records.Add "LFS|8.4 Mitigation & Control|D82|A1203A|A12031|IntegerValue|B82|1708997244"
    records.Add "LFT|8.4 Mitigation & Control|D83|A1203A|A12031|IntegerValue|B83|380525302"
    records.Add "LFU|8.4 Mitigation & Control|C88|A1203A|A12031|EnumerationCDD|B88|1281383680"
    records.Add "LFV|8.4 Mitigation & Control|C89|A1203A|A12031|EnumerationCDD|B89|1843301630"
    records.Add "LFW|8.4 Mitigation & Control|C90|A1203A|A12031|EnumerationCDD|B90|239872886"
    records.Add "LFX|8.4 Mitigation & Control|C91|A1203A|A12031|EnumerationCDD|B91|290477190"
    records.Add "LFY|8.4 Mitigation & Control|C92|A1203A|A12031|EnumerationCDD|B92|2048635571"
    records.Add "LFZ|8.4 Mitigation & Control|C93|A1203A|A12031|EnumerationCDD|B93|513869623"
    records.Add "LGA|8.4 Mitigation & Control|D88|A1203A|A12031|EnumerationCDD|B88|1281383680"
    records.Add "LGB|8.4 Mitigation & Control|D89|A1203A|A12031|EnumerationCDD|B89|1843301630"
    records.Add "LGC|8.4 Mitigation & Control|D90|A1203A|A12031|EnumerationCDD|B90|239872886"
    records.Add "LGD|8.4 Mitigation & Control|D91|A1203A|A12031|EnumerationCDD|B91|290477190"
    records.Add "LGE|8.4 Mitigation & Control|D92|A1203A|A12031|EnumerationCDD|B92|2048635571"
    records.Add "LGF|8.4 Mitigation & Control|D93|A1203A|A12031|EnumerationCDD|B93|513869623"
    records.Add "LGG|8.4 Mitigation & Control|C96|A1203A|A12031|EnumerationYesNo|B96|177461107"
    records.Add "LGH|8.4 Mitigation & Control|D96|A1203A|A12031|EnumerationYesNo|B96|177461107"
    records.Add "LGI|8.4 Mitigation & Control|C98|A1203A|A12031|EnumerationFrequency|B98|1393704543"
    records.Add "LGJ|8.4 Mitigation & Control|C99|A1203A|A12031|EnumerationFrequency|B99|2049025752"
    records.Add "LGK|8.4 Mitigation & Control|C100|A1203A|A12031|EnumerationFrequency|B100|1978708663"
    records.Add "LGL|8.4 Mitigation & Control|C101|A1203A|A12031|EnumerationFrequency|B101|2452184"
    records.Add "LGM|8.4 Mitigation & Control|C102|A1203A|A12031|EnumerationFrequency|B102|1390031435"
    records.Add "LGN|8.4 Mitigation & Control|D98|A1203A|A12031|EnumerationFrequency|B98|1393704543"
    records.Add "LGO|8.4 Mitigation & Control|D99|A1203A|A12031|EnumerationFrequency|B99|2049025752"
    records.Add "LGP|8.4 Mitigation & Control|D100|A1203A|A12031|EnumerationFrequency|B100|1978708663"
    records.Add "LGQ|8.4 Mitigation & Control|D101|A1203A|A12031|EnumerationFrequency|B101|2452184"
    records.Add "LGR|8.4 Mitigation & Control|D102|A1203A|A12031|EnumerationFrequency|B102|1390031435"
    records.Add "LGS|8.4 Mitigation & Control|C104|A1203A|A12031|DateValue|B104|1355524720"
    records.Add "LGT|8.4 Mitigation & Control|D104|A1203A|A12031|DateValue|B104|1355524720"
    records.Add "LGU|8.4 Mitigation & Control|C107|A1203A|A12031|IntegerValue|B107|725850943"
    records.Add "LGV|8.4 Mitigation & Control|C108|A1203A|A12031|IntegerValue|B108|553687269"
    records.Add "LGW|8.4 Mitigation & Control|C109|A1203A|A12031|IntegerValue|B109|886586389"
    records.Add "LGX|8.4 Mitigation & Control|C110|A1203A|A12031|IntegerValue|B110|1928692989"
    records.Add "LGY|8.4 Mitigation & Control|D107|A1203A|A12031|IntegerValue|B107|725850943"
    records.Add "LGZ|8.4 Mitigation & Control|D108|A1203A|A12031|IntegerValue|B108|553687269"
    records.Add "LHA|8.4 Mitigation & Control|D109|A1203A|A12031|IntegerValue|B109|886586389"
    records.Add "LHB|8.4 Mitigation & Control|D110|A1203A|A12031|IntegerValue|B110|1928692989"
    records.Add "LHC|8.4 Mitigation & Control|C113|A1203A|A12031|IntegerValue|B113|757985860"
    records.Add "LHD|8.4 Mitigation & Control|C114|A1203A|A12031|IntegerValue|B114|460233053"
    records.Add "LHE|8.4 Mitigation & Control|D113|A1203A|A12031|IntegerValue|B113|757985860"
    records.Add "LHF|8.4 Mitigation & Control|D114|A1203A|A12031|IntegerValue|B114|460233053"
    records.Add "LHG|8.4 Mitigation & Control|C117|A1203A|A12031|DecimalValue|B117|425020038"
    records.Add "LHH|8.4 Mitigation & Control|C118|A1203A|A12031|DecimalValue|B118|394844484"
    records.Add "LHI|8.4 Mitigation & Control|C119|A1203A|A12031|DecimalValue|B119|2452184"
    records.Add "LHJ|8.4 Mitigation & Control|C120|A1203A|A12031|DecimalValue|B120|420924238"
    records.Add "LHK|8.4 Mitigation & Control|D117|A1203A|A12031|DecimalValue|B117|425020038"
    records.Add "LHL|8.4 Mitigation & Control|D118|A1203A|A12031|DecimalValue|B118|394844484"
    records.Add "LHM|8.4 Mitigation & Control|D119|A1203A|A12031|DecimalValue|B119|2452184"
    records.Add "LHN|8.4 Mitigation & Control|D120|A1203A|A12031|DecimalValue|B120|420924238"
    records.Add "LHO|8.4 Mitigation & Control|C122|A1203A|A12031|EnumerationYesNoNA|B122|792731794"
    records.Add "LHP|8.4 Mitigation & Control|D122|A1203A|A12031|EnumerationYesNoNA|B122|792731794"
    records.Add "LHQ|8.4 Mitigation & Control|C125|A1203A|A12031|PercentageValue|B125|912644999"
    records.Add "LHR|8.4 Mitigation & Control|C126|A1203A|A12031|PercentageValue|B126|1978708663"
    records.Add "LHS|8.4 Mitigation & Control|C127|A1203A|A12031|PercentageValue|B127|2452184"
    records.Add "LHT|8.4 Mitigation & Control|D125|A1203A|A12031|PercentageValue|B125|912644999"
End Sub

Private Sub MapPart3(ByRef records As Collection)
    records.Add "LHU|8.4 Mitigation & Control|D126|A1203A|A12031|PercentageValue|B126|1978708663"
    records.Add "LHV|8.4 Mitigation & Control|D127|A1203A|A12031|PercentageValue|B127|2452184"
    records.Add "LHW|8.4 Mitigation & Control|C130|A1203A|A12031|IntegerValue|B130|700766425"
    records.Add "LHX|8.4 Mitigation & Control|C131|A1203A|A12031|IntegerValue|B131|1628086517"
    records.Add "LHY|8.4 Mitigation & Control|C132|A1203A|A12031|IntegerValue|B132|601507047"
    records.Add "LHZ|8.4 Mitigation & Control|C133|A1203A|A12031|IntegerValue|B133|270354697"
    records.Add "LIA|8.4 Mitigation & Control|D132|A1203A|A12031|IntegerValue|B132|601507047"
    records.Add "LIB|8.4 Mitigation & Control|D133|A1203A|A12031|IntegerValue|B133|270354697"
    records.Add "LIC|8.4 Mitigation & Control|C138|A1203A|A12031|EnumerationYesNoNA|B138|1981177934"
    records.Add "LID|8.4 Mitigation & Control|C139|A1203A|A12031|EnumerationFrequency|B139|1856008064"
    records.Add "LIE|8.4 Mitigation & Control|D138|A1203A|A12031|EnumerationYesNoNA|B138|1981177934"
    records.Add "LIF|8.4 Mitigation & Control|D139|A1203A|A12031|EnumerationFrequency|B139|1856008064"
    records.Add "LIG|8.4 Mitigation & Control|C142|A1203A|A12031|EnumerationYesNoNA|B142|2016637792"
    records.Add "LIH|8.4 Mitigation & Control|C143|A1203A|A12031|EnumerationYesNoNA|B143|1404609399"
    records.Add "LII|8.4 Mitigation & Control|C144|A1203A|A12031|EnumerationYesNoNA|B144|74195260"
    records.Add "LIJ|8.4 Mitigation & Control|C145|A1203A|A12031|EnumerationYesNoNA|B145|2027301300"
    records.Add "LIK|8.4 Mitigation & Control|C146|A1203A|A12031|EnumerationYesNoNA|B146|2110073219"
    records.Add "LIL|8.4 Mitigation & Control|C147|A1203A|A12031|EnumerationYesNoNA|B147|295033668"
    records.Add "LIM|8.4 Mitigation & Control|C148|A1203A|A12031|EnumerationYesNoNA|B148|1631194507"
    records.Add "LIN|8.4 Mitigation & Control|D142|A1203A|A12031|EnumerationYesNoNA|B142|2016637792"
    records.Add "LIO|8.4 Mitigation & Control|D143|A1203A|A12031|EnumerationYesNoNA|B143|1404609399"
    records.Add "LIP|8.4 Mitigation & Control|D144|A1203A|A12031|EnumerationYesNoNA|B144|74195260"
    records.Add "LIQ|8.4 Mitigation & Control|D145|A1203A|A12031|EnumerationYesNoNA|B145|2027301300"
    records.Add "LIR|8.4 Mitigation & Control|D146|A1203A|A12031|EnumerationYesNoNA|B146|2110073219"
    records.Add "LIS|8.4 Mitigation & Control|D147|A1203A|A12031|EnumerationYesNoNA|B147|295033668"
    records.Add "LIT|8.4 Mitigation & Control|D148|A1203A|A12031|EnumerationYesNoNA|B148|1631194507"
    records.Add "LIU|8.4 Mitigation & Control|C150|A1203A|A12031|EnumerationMonitoringApproach|B150|604731773"
    records.Add "LIV|8.4 Mitigation & Control|C151|A1203A|A12031|EnumerationModelValidation|B151|665786603"
    records.Add "LIW|8.4 Mitigation & Control|C152|A1203A|A12031|EnumerationYesNoNA|B152|2008171"
    records.Add "LIX|8.4 Mitigation & Control|D150|A1203A|A12031|EnumerationYesNoNA|B150|604731773"
    records.Add "LIY|8.4 Mitigation & Control|D151|A1203A|A12031|EnumerationMonitoringApproach|B151|665786603"
    records.Add "LIZ|8.4 Mitigation & Control|D152|A1203A|A12031|EnumerationModelValidation|B152|2008171"
    records.Add "LJA|8.4 Mitigation & Control|C154|A1203A|A12031|EnumerationYesNoNA|B154|1588499589"
    records.Add "LJB|8.4 Mitigation & Control|D154|A1203A|A12031|EnumerationYesNoNA|B154|1588499589"
    records.Add "LJC|8.4 Mitigation & Control|C157|A1203A|A12031|EnumerationYesNoNA|B157|2020808454"
    records.Add "LJD|8.4 Mitigation & Control|C158|A1203A|A12031|EnumerationYesNoNA|B158|59733447"
    records.Add "LJE|8.4 Mitigation & Control|C159|A1203A|A12031|EnumerationYesNoNA|B159|1022426218"
    records.Add "LJF|8.4 Mitigation & Control|C160|A1203A|A12031|EnumerationYesNoNA|B160|1487416784"
    records.Add "LJG|8.4 Mitigation & Control|C161|A1203A|A12031|EnumerationYesNoNA|B161|76017736"
    records.Add "LJH|8.4 Mitigation & Control|C162|A1203A|A12031|EnumerationYesNoNA|B162|2118087349"
    records.Add "LJI|8.4 Mitigation & Control|C163|A1203A|A12031|EnumerationYesNoNA|B163|76517104"
    records.Add "LJJ|8.4 Mitigation & Control|C164|A1203A|A12031|String500|B164|2122788987"
    records.Add "LJK|8.4 Mitigation & Control|D157|A1203A|A12031|EnumerationYesNoNA|B157|2020808454"
    records.Add "LJL|8.4 Mitigation & Control|D158|A1203A|A12031|EnumerationYesNoNA|B158|59733447"
    records.Add "LJM|8.4 Mitigation & Control|D161|A1203A|A12031|EnumerationYesNoNA|B161|76017736"
    records.Add "LJN|8.4 Mitigation & Control|D162|A1203A|A12031|EnumerationYesNoNA|B162|2118087349"
    records.Add "LJO|8.4 Mitigation & Control|D163|A1203A|A12031|EnumerationYesNoNA|B163|76517104"
    records.Add "LJP|8.4 Mitigation & Control|D164|A1203A|A12031|String500|B164|2122788987"
    records.Add "LJQ|8.4 Mitigation & Control|C166|A1203A|A12031|EnumerationYesNoNA|B166|565730629"
    records.Add "LJR|8.4 Mitigation & Control|C167|A1203A|A12031|DateValue|B167|1307421599"
    records.Add "LJS|8.4 Mitigation & Control|C168|A1203A|A12031|EnumerationYesNoNA|B168|2045941472"
    records.Add "LJT|8.4 Mitigation & Control|C169|A1203A|A12031|EnumerationDetectionTime|B169|1683526163"
    records.Add "LJU|8.4 Mitigation & Control|D166|A1203A|A12031|EnumerationYesNoNA|B166|565730629"
    records.Add "LJV|8.4 Mitigation & Control|D167|A1203A|A12031|DateValue|B167|1307421599"
    records.Add "LJW|8.4 Mitigation & Control|D168|A1203A|A12031|EnumerationYesNoNA|B168|2045941472"
    records.Add "LJX|8.4 Mitigation & Control|D169|A1203A|A12031|EnumerationDetectionTime|B169|1683526163"
    records.Add "LJY|8.4 Mitigation & Control|C172|A1203A|A12031|IntegerValue|B172|1387750479"
    records.Add "LJZ|8.4 Mitigation & Control|C177|A1203A|A12031|IntegerValue|B177|195422839"
    records.Add "LKA|8.4 Mitigation & Control|C178|A1203A|A12031|IntegerValue|B178|550354251"
    records.Add "LKB|8.4 Mitigation & Control|C179|A1203A|A12031|PercentageValue|B179|1133043437"
End Sub

Private Sub MapPart4(ByRef records As Collection)
    records.Add "LKC|8.4 Mitigation & Control|C180|A1203A|A12031|IntegerValue|B180|1815380415"
    records.Add "LKD|8.4 Mitigation & Control|C181|A1203A|A12031|IntegerValue|B181|2032185690"
    records.Add "LKE|8.4 Mitigation & Control|C182|A1203A|A12031|IntegerValue|B182|1625706833"
    records.Add "LKF|8.4 Mitigation & Control|C183|A1203A|A12031|IntegerValue|B183|1059540111"
    records.Add "LKG|8.4 Mitigation & Control|C184|A1203A|A12031|IntegerValue|B184|2032742661"
    records.Add "LKH|8.4 Mitigation & Control|C185|A1203A|A12031|DecimalValue|B185|1766499786"
    records.Add "LKI|8.4 Mitigation & Control|C186|A1203A|A12031|IntegerValue|B186|1443577462"
    records.Add "LKJ|8.4 Mitigation & Control|C187|A1203A|A12031|DecimalValue|B187|936311529"
    records.Add "LKK|8.4 Mitigation & Control|D177|A1203A|A12031|IntegerValue|B177|195422839"
    records.Add "LKL|8.4 Mitigation & Control|D178|A1203A|A12031|IntegerValue|B178|550354251"
    records.Add "LKM|8.4 Mitigation & Control|D179|A1203A|A12031|PercentageValue|B179|1133043437"
    records.Add "LKN|8.4 Mitigation & Control|D180|A1203A|A12031|IntegerValue|B180|1815380415"
    records.Add "LKO|8.4 Mitigation & Control|D181|A1203A|A12031|IntegerValue|B181|2032185690"
    records.Add "LKP|8.4 Mitigation & Control|D182|A1203A|A12031|IntegerValue|B182|1625706833"
    records.Add "LKQ|8.4 Mitigation & Control|D183|A1203A|A12031|IntegerValue|B183|1059540111"
    records.Add "LKR|8.4 Mitigation & Control|D184|A1203A|A12031|IntegerValue|B184|2032742661"
    records.Add "LKS|8.4 Mitigation & Control|D185|A1203A|A12031|DecimalValue|B185|1766499786"
    records.Add "LKT|8.4 Mitigation & Control|D186|A1203A|A12031|IntegerValue|B186|1443577462"
    records.Add "LKU|8.4 Mitigation & Control|D187|A1203A|A12031|DecimalValue|B187|936311529"
    records.Add "LKV|8.4 Mitigation & Control|C191|A1203A|A12031|EnumerationYesNo|B191|1755416898"
    records.Add "LKW|8.4 Mitigation & Control|C195|A1203A|A12031|IntegerValue|B195|1549319147"
    records.Add "LKX|8.4 Mitigation & Control|C196|A1203A|A12031|IntegerValue|B196|973217486"
    records.Add "LKY|8.4 Mitigation & Control|C197|A1203A|A12031|IntegerValue|B197|2014993732"
    records.Add "LKZ|8.4 Mitigation & Control|C198|A1203A|A12031|IntegerValue|B198|747774027"
    records.Add "LLA|8.4 Mitigation & Control|C199|A1203A|A12031|IntegerValue|B199|1980260071"
    records.Add "LLB|8.4 Mitigation & Control|C200|A1203A|A12031|IntegerValue|B200|1899793439"
    records.Add "LLC|8.4 Mitigation & Control|C201|A1203A|A12031|IntegerValue|B201|207694753"
    records.Add "LLD|8.4 Mitigation & Control|C202|A1203A|A12031|IntegerValue|B202|1720426881"
    records.Add "LLE|8.4 Mitigation & Control|C203|A1203A|A12031|IntegerValue|B203|585428131"
    records.Add "LLF|8.4 Mitigation & Control|C204|A1203A|A12031|IntegerValue|B204|424122303"
    records.Add "LLG|8.4 Mitigation & Control|C205|A1203A|A12031|IntegerValue|B205|1791543088"
    records.Add "LLH|8.4 Mitigation & Control|C206|A1203A|A12031|IntegerValue|B206|382882081"
    records.Add "LLI|8.4 Mitigation & Control|C207|A1203A|A12031|IntegerValue|B207|1167655784"
    records.Add "LLJ|8.4 Mitigation & Control|C208|A1203A|A12031|IntegerValue|B208|449738432"
    records.Add "LLK|8.4 Mitigation & Control|D195|A1203A|A12031|IntegerValue|B195|1549319147"
    records.Add "LLL|8.4 Mitigation & Control|D196|A1203A|A12031|IntegerValue|B196|973217486"
    records.Add "LLM|8.4 Mitigation & Control|D197|A1203A|A12031|IntegerValue|B197|2014993732"
    records.Add "LLN|8.4 Mitigation & Control|D198|A1203A|A12031|IntegerValue|B198|747774027"
    records.Add "LLO|8.4 Mitigation & Control|D199|A1203A|A12031|IntegerValue|B199|1980260071"
    records.Add "LLP|8.4 Mitigation & Control|D200|A1203A|A12031|IntegerValue|B200|1899793439"
    records.Add "LLQ|8.4 Mitigation & Control|D201|A1203A|A12031|IntegerValue|B201|207694753"
    records.Add "LLR|8.4 Mitigation & Control|D202|A1203A|A12031|IntegerValue|B202|1720426881"
    records.Add "LLS|8.4 Mitigation & Control|D203|A1203A|A12031|IntegerValue|B203|585428131"
    records.Add "LLT|8.4 Mitigation & Control|D204|A1203A|A12031|IntegerValue|B204|424122303"
    records.Add "LLU|8.4 Mitigation & Control|D205|A1203A|A12031|IntegerValue|B205|1791543088"
    records.Add "LLV|8.4 Mitigation & Control|D206|A1203A|A12031|IntegerValue|B206|382882081"
    records.Add "LLW|8.4 Mitigation & Control|D207|A1203A|A12031|IntegerValue|B207|1167655784"
    records.Add "LLX|8.4 Mitigation & Control|D208|A1203A|A12031|IntegerValue|B208|449738432"
    records.Add "LLY|8.4 Mitigation & Control|C211|A1203A|A12031|DecimalValue|B211|1597025068"
    records.Add "LLZ|8.4 Mitigation & Control|C215|A1203A|A12031|EnumerationYesNo|B215|1413692325"
    records.Add "LMA|8.4 Mitigation & Control|C216|A1203A|A12031|DecimalValue|B216|1772583534"
    records.Add "LMB|8.4 Mitigation & Control|D215|A1203A|A12031|EnumerationYesNo|B215|1413692325"
    records.Add "LMC|8.4 Mitigation & Control|D216|A1203A|A12031|DecimalValue|B216|1772583534"
    records.Add "LMD|8.4 Mitigation & Control|C219|A1203A|A12031|EnumerationYesNo|B219|533925649"
    records.Add "LME|8.4 Mitigation & Control|C220|A1203A|A12031|EnumerationYesNo|B220|809880783"
    records.Add "LMF|8.4 Mitigation & Control|C221|A1203A|A12031|EnumerationYesNo|B221|1487433647"
    records.Add "LMG|8.4 Mitigation & Control|C222|A1203A|A12031|EnumerationYesNo|B222|76517104"
    records.Add "LMH|8.4 Mitigation & Control|C223|A1203A|A12031|String500|B223|1661430485"
    records.Add "LMI|8.4 Mitigation & Control|D219|A1203A|A12031|EnumerationYesNo|B219|533925649"
    records.Add "LMJ|8.4 Mitigation & Control|D220|A1203A|A12031|EnumerationYesNo|B220|809880783"
End Sub

Private Sub MapPart5(ByRef records As Collection)
    records.Add "LMK|8.4 Mitigation & Control|D221|A1203A|A12031|EnumerationYesNo|B221|1487433647"
    records.Add "LML|8.4 Mitigation & Control|D222|A1203A|A12031|EnumerationYesNo|B222|76517104"
    records.Add "LMM|8.4 Mitigation & Control|D223|A1203A|A12031|String500|B223|1661430485"
    records.Add "LMN|8.4 Mitigation & Control|C226|A1203A|A12031|EnumerationYesNo|B226|1273487561"
    records.Add "LMO|8.4 Mitigation & Control|C227|A1203A|A12031|EnumerationYesNo|B227|1950000346"
    records.Add "LMP|8.4 Mitigation & Control|D226|A1203A|A12031|EnumerationYesNo|B226|1273487561"
    records.Add "LMQ|8.4 Mitigation & Control|D227|A1203A|A12031|EnumerationYesNo|B227|1950000346"
    records.Add "LMR|8.4 Mitigation & Control|C230|A1203A|A12031|IntegerValue|B230|1072961842"
    records.Add "LMS|8.4 Mitigation & Control|C231|A1203A|A12031|IntegerValue|B231|791676615"
    records.Add "LMT|8.4 Mitigation & Control|C232|A1203A|A12031|IntegerValue|B232|2033000994"
    records.Add "LMU|8.4 Mitigation & Control|D230|A1203A|A12031|IntegerValue|B230|1072961842"
    records.Add "LMV|8.4 Mitigation & Control|D231|A1203A|A12031|IntegerValue|B231|791676615"
    records.Add "LMW|8.4 Mitigation & Control|D232|A1203A|A12031|IntegerValue|B232|2033000994"
    records.Add "LMX|8.4 Mitigation & Control|C235|A1203A|A12031|IntegerValue|B235|65718865"
    records.Add "LMY|8.4 Mitigation & Control|C236|A1203A|A12031|DecimalValue|B236|1747602265"
    records.Add "LMZ|8.4 Mitigation & Control|C241|A1203A|A12031|EnumerationOutsource|B241|24731827"
    records.Add "LNA|8.4 Mitigation & Control|C242|A1203A|A12031|EnumerationOutsource|B242|235263333"
    records.Add "LNB|8.4 Mitigation & Control|C243|A1203A|A12031|EnumerationOutsource|B243|880441845"
    records.Add "LNC|8.4 Mitigation & Control|C244|A1203A|A12031|EnumerationOutsource|B244|1306640399"
    records.Add "LND|8.4 Mitigation & Control|C245|A1203A|A12031|EnumerationOutsource|B245|1424215748"
    records.Add "LNE|8.4 Mitigation & Control|C246|A1203A|A12031|EnumerationOutsource|B246|1287919970"
    records.Add "LNF|8.4 Mitigation & Control|C247|A1203A|A12031|EnumerationOutsource|B247|447501236"
    records.Add "LNG|8.4 Mitigation & Control|C248|A1203A|A12031|EnumerationOutsource|B248|1340764510"
    records.Add "LNH|8.4 Mitigation & Control|C249|A1203A|A12031|EnumerationOutsource|B249|1011577919"
    records.Add "LNI|8.4 Mitigation & Control|C250|A1203A|A12031|EnumerationOutsource|B250|18422217"
    records.Add "LNJ|8.4 Mitigation & Control|C251|A1203A|A12031|EnumerationOutsource|B251|729709613"
    records.Add "LNK|8.4 Mitigation & Control|C252|A1203A|A12031|String500|B252|1290477889"
    records.Add "LNL|8.4 Mitigation & Control|D241|A1203A|A12031|EnumerationOutsource|B241|24731827"
    records.Add "LNM|8.4 Mitigation & Control|D242|A1203A|A12031|EnumerationOutsource|B242|235263333"
    records.Add "LNN|8.4 Mitigation & Control|D243|A1203A|A12031|EnumerationOutsource|B243|880441845"
    records.Add "LNO|8.4 Mitigation & Control|D244|A1203A|A12031|EnumerationOutsource|B244|1306640399"
    records.Add "LNP|8.4 Mitigation & Control|D245|A1203A|A12031|EnumerationOutsource|B245|1424215748"
    records.Add "LNQ|8.4 Mitigation & Control|D246|A1203A|A12031|EnumerationOutsource|B246|1287919970"
    records.Add "LNR|8.4 Mitigation & Control|D247|A1203A|A12031|EnumerationOutsource|B247|447501236"
    records.Add "LNS|8.4 Mitigation & Control|D248|A1203A|A12031|EnumerationOutsource|B248|1340764510"
    records.Add "LNT|8.4 Mitigation & Control|D249|A1203A|A12031|EnumerationOutsource|B249|1011577919"
    records.Add "LNU|8.4 Mitigation & Control|D250|A1203A|A12031|EnumerationOutsource|B250|18422217"
    records.Add "LNV|8.4 Mitigation & Control|D251|A1203A|A12031|EnumerationOutsource|B251|729709613"
    records.Add "LNW|8.4 Mitigation & Control|D252|A1203A|A12031|String500|B252|1290477889"
    records.Add "LNX|8.4 Mitigation & Control|C255|A1203A|A12031|PercentageValue|B255|1902530903"
    records.Add "LNY|8.4 Mitigation & Control|C256|A1203A|A12031|EnumerationYesNoNA|B256|1211143384"
    records.Add "LNZ|8.4 Mitigation & Control|D255|A1203A|A12031|PercentageValue|B255|1902530903"
    records.Add "LOA|8.4 Mitigation & Control|D256|A1203A|A12031|EnumerationYesNoNA|B256|1211143384"
    records.Add "LOB|8.4 Mitigation & Control|C258|A1203A|A12031|EnumerationYesNoNA|B258|390841210"
    records.Add "LOC|8.4 Mitigation & Control|C259|A1203A|A12031|EnumerationYesNoNA|B259|570253193"
    records.Add "LOD|8.4 Mitigation & Control|C260|A1203A|A12031|EnumerationYesNoNA|B260|2055947971"
    records.Add "LOE|8.4 Mitigation & Control|C261|A1203A|A12031|EnumerationYesNoNA|B261|1802528110"
    records.Add "LOF|8.4 Mitigation & Control|C262|A1203A|A12031|String500|B262|1404835929"
    records.Add "LOG|8.4 Mitigation & Control|D258|A1203A|A12031|EnumerationYesNoNA|B258|390841210"
    records.Add "LOH|8.4 Mitigation & Control|D259|A1203A|A12031|EnumerationYesNoNA|B259|570253193"
    records.Add "LOI|8.4 Mitigation & Control|D260|A1203A|A12031|EnumerationYesNoNA|B260|2055947971"
    records.Add "LOJ|8.4 Mitigation & Control|D261|A1203A|A12031|EnumerationYesNoNA|B261|1802528110"
    records.Add "LOK|8.4 Mitigation & Control|D262|A1203A|A12031|String500|B262|1404835929"
    records.Add "LOL|8.4 Mitigation & Control|C265|A1203A|A12031|EnumerationYesNoNA|B265|1396634375"
    records.Add "LOM|8.4 Mitigation & Control|C266|A1203A|A12031|EnumerationYesNoNA|B266|1142570223"
    records.Add "LON|8.4 Mitigation & Control|D265|A1203A|A12031|EnumerationYesNoNA|B265|1396634375"
    records.Add "LOO|8.4 Mitigation & Control|D266|A1203A|A12031|EnumerationYesNoNA|B266|1142570223"
    records.Add "LOP|8.4 Mitigation & Control|C271|A1203A|A12031|PercentageValue|B271|531221994"
    records.Add "LOQ|8.4 Mitigation & Control|C272|A1203A|A12031|PercentageValue|B272|1454833862"
    records.Add "LOR|8.4 Mitigation & Control|C273|A1203A|A12031|PercentageValue|B273|1596642414"
End Sub

Private Sub MapPart6(ByRef records As Collection)
    records.Add "LOS|8.4 Mitigation & Control|C274|A1203A|A12031|PercentageValue|B274|1368848164"
    records.Add "LOT|8.4 Mitigation & Control|C275|A1203A|A12031|PercentageValue|B275|76517104"
    records.Add "LOU|8.4 Mitigation & Control|D271|A1203A|A12031|PercentageValue|B271|531221994"
    records.Add "LOV|8.4 Mitigation & Control|D272|A1203A|A12031|PercentageValue|B272|1454833862"
    records.Add "LOW|8.4 Mitigation & Control|D273|A1203A|A12031|PercentageValue|B273|1596642414"
    records.Add "LOX|8.4 Mitigation & Control|D274|A1203A|A12031|PercentageValue|B274|1368848164"
    records.Add "LOY|8.4 Mitigation & Control|D275|A1203A|A12031|PercentageValue|B275|76517104"
    records.Add "LOZ|8.4 Mitigation & Control|C278|A1203A|A12031|IntegerValue|B278|565803125"
    records.Add "LPA|8.4 Mitigation & Control|C283|A1203A|A12031|DateValue|B283|816777597"
    records.Add "LPB|8.4 Mitigation & Control|C284|A1203A|A12031|DateValue|B284|782924878"
    records.Add "LPC|8.4 Mitigation & Control|C285|A1203A|A12031|DateValue|B285|879047402"
    records.Add "LPD|8.4 Mitigation & Control|C286|A1203A|A12031|DateValue|B286|559241562"
    records.Add "LPE|8.4 Mitigation & Control|C287|A1203A|A12031|DateValue|B287|746252736"
    records.Add "LPF|8.4 Mitigation & Control|C288|A1203A|A12031|DateValue|B288|186464138"
    records.Add "LPG|8.4 Mitigation & Control|C289|A1203A|A12031|DateValue|B289|2093448417"
    records.Add "LPH|8.4 Mitigation & Control|C290|A1203A|A12031|DateValue|B290|556452828"
    records.Add "LPI|8.4 Mitigation & Control|C291|A1203A|A12031|DateValue|B291|1881693956"
    records.Add "LPJ|8.4 Mitigation & Control|C292|A1203A|A12031|DateValue|B292|5096075"
    records.Add "LPK|8.4 Mitigation & Control|C293|A1203A|A12031|DateValue|B293|1446035469"
    records.Add "LPL|8.4 Mitigation & Control|C294|A1203A|A12031|EnumerationYesNo|B294|818245780"
    records.Add "LPM|8.4 Mitigation & Control|C295|A1203A|A12031|String500|B295|1764957583"
    records.Add "LPN|8.4 Mitigation & Control|D283|A1203A|A12031|String500|B283|816777597"
    records.Add "LPO|8.4 Mitigation & Control|D284|A1203A|A12031|String500|B284|782924878"
    records.Add "LPP|8.4 Mitigation & Control|D285|A1203A|A12031|String500|B285|879047402"
    records.Add "LPQ|8.4 Mitigation & Control|D286|A1203A|A12031|String500|B286|559241562"
    records.Add "LPR|8.4 Mitigation & Control|D287|A1203A|A12031|String500|B287|746252736"
    records.Add "LPS|8.4 Mitigation & Control|D288|A1203A|A12031|String500|B288|186464138"
    records.Add "LPT|8.4 Mitigation & Control|D289|A1203A|A12031|String500|B289|2093448417"
    records.Add "LPU|8.4 Mitigation & Control|D290|A1203A|A12031|String500|B290|556452828"
    records.Add "LPV|8.4 Mitigation & Control|D291|A1203A|A12031|String500|B291|1881693956"
    records.Add "LPW|8.4 Mitigation & Control|D292|A1203A|A12031|String500|B292|5096075"
    records.Add "LPX|8.4 Mitigation & Control|D293|A1203A|A12031|String500|B293|1446035469"
    records.Add "LPY|8.4 Mitigation & Control|D295|A1203A|A12031|String500|B295|1764957583"
    records.Add "LPZ|8.4 Mitigation & Control|C298|A1203A|A12031|EnumerationFrequency|B298|1872198132"
    records.Add "LQA|8.4 Mitigation & Control|D298|A1203A|A12031|EnumerationFrequency|B298|1872198132"
    records.Add "LQB|8.4 Mitigation & Control|C300|A1203A|A12031|IntegerValue|B300|720250472"
    records.Add "LQC|8.4 Mitigation & Control|C306|A1203A|A12031|DateValue|B306|816777597"
    records.Add "LQD|8.4 Mitigation & Control|C307|A1203A|A12031|DateValue|B307|782924878"
    records.Add "LQE|8.4 Mitigation & Control|C308|A1203A|A12031|DateValue|B308|879047402"
    records.Add "LQF|8.4 Mitigation & Control|C309|A1203A|A12031|DateValue|B309|559241562"
    records.Add "LQG|8.4 Mitigation & Control|C310|A1203A|A12031|DateValue|B310|746252736"
    records.Add "LQH|8.4 Mitigation & Control|C311|A1203A|A12031|DateValue|B311|186464138"
    records.Add "LQI|8.4 Mitigation & Control|C312|A1203A|A12031|DateValue|B312|777569128"
    records.Add "LQJ|8.4 Mitigation & Control|C313|A1203A|A12031|DateValue|B313|556452828"
    records.Add "LQK|8.4 Mitigation & Control|C314|A1203A|A12031|DateValue|B314|1881693956"
    records.Add "LQL|8.4 Mitigation & Control|C315|A1203A|A12031|DateValue|B315|5096075"
    records.Add "LQM|8.4 Mitigation & Control|C316|A1203A|A12031|DateValue|B316|1446035469"
    records.Add "LQN|8.4 Mitigation & Control|C317|A1203A|A12031|EnumerationYesNo|B317|818245780"
    records.Add "LQO|8.4 Mitigation & Control|C318|A1203A|A12031|String500|B318|1764957583"
    records.Add "LQP|8.4 Mitigation & Control|D306|A1203A|A12031|String500|B306|816777597"
    records.Add "LQQ|8.4 Mitigation & Control|D307|A1203A|A12031|String500|B307|782924878"
    records.Add "LQR|8.4 Mitigation & Control|D308|A1203A|A12031|String500|B308|879047402"
    records.Add "LQS|8.4 Mitigation & Control|D309|A1203A|A12031|String500|B309|559241562"
    records.Add "LQT|8.4 Mitigation & Control|D310|A1203A|A12031|String500|B310|746252736"
    records.Add "LQU|8.4 Mitigation & Control|D311|A1203A|A12031|String500|B311|186464138"
    records.Add "LQV|8.4 Mitigation & Control|D312|A1203A|A12031|String500|B312|777569128"
    records.Add "LQW|8.4 Mitigation & Control|D313|A1203A|A12031|String500|B313|556452828"
    records.Add "LQX|8.4 Mitigation & Control|D314|A1203A|A12031|String500|B314|1881693956"
    records.Add "LQY|8.4 Mitigation & Control|D315|A1203A|A12031|String500|B315|5096075"
    records.Add "LQZ|8.4 Mitigation & Control|D316|A1203A|A12031|String500|B316|1446035469"
End Sub

Private Sub MapPart7(ByRef records As Collection)
    records.Add "LRA|8.4 Mitigation & Control|D318|A1203A|A12031|String500|B318|1764957583"
    records.Add "LRB|8.4 Mitigation & Control|C321|A1203A|A12031|EnumerationFrequency|B321|805054581"
    records.Add "LRC|8.4 Mitigation & Control|D321|A1203A|A12031|EnumerationFrequency|B321|805054581"
    records.Add "LRD|8.4 Mitigation & Control|C324|A1203A|A12031|EnumerationYesNo|B324|1899793439"
    records.Add "LRE|8.4 Mitigation & Control|C325|A1203A|A12031|EnumerationYesNo|B325|1837590984"
    records.Add "LRF|8.4 Mitigation & Control|C326|A1203A|A12031|EnumerationYesNo|B326|537424550"
    records.Add "LRG|8.4 Mitigation & Control|C327|A1203A|A12031|EnumerationYesNo|B327|17478703"
    records.Add "LRH|8.4 Mitigation & Control|C328|A1203A|A12031|EnumerationYesNo|B328|76517104"
    records.Add "LRI|8.4 Mitigation & Control|C329|A1203A|A12031|String500|B329|1026593944"
    records.Add "LRJ|8.4 Mitigation & Control|D324|A1203A|A12031|String500|B324|1899793439"
    records.Add "LRK|8.4 Mitigation & Control|D325|A1203A|A12031|String500|B325|1837590984"
    records.Add "LRL|8.4 Mitigation & Control|D326|A1203A|A12031|String500|B326|537424550"
    records.Add "LRM|8.4 Mitigation & Control|D327|A1203A|A12031|String500|B327|17478703"
    records.Add "LRN|8.4 Mitigation & Control|D329|A1203A|A12031|String500|B329|1026593944"
    records.Add "LRO|8.4 Mitigation & Control|C335|A1203A|A12031|EnumerationFrequency|B335|1899793439"
    records.Add "LRP|8.4 Mitigation & Control|C336|A1203A|A12031|EnumerationFrequency|B336|1837590984"
    records.Add "LRQ|8.4 Mitigation & Control|C337|A1203A|A12031|EnumerationFrequency|B337|537424550"
    records.Add "LRR|8.4 Mitigation & Control|C338|A1203A|A12031|EnumerationFrequency|B338|17478703"
    records.Add "LRS|8.4 Mitigation & Control|D335|A1203A|A12031|EnumerationFrequency|B335|1899793439"
    records.Add "LRT|8.4 Mitigation & Control|D336|A1203A|A12031|EnumerationFrequency|B336|1837590984"
    records.Add "LRU|8.4 Mitigation & Control|D337|A1203A|A12031|EnumerationFrequency|B337|537424550"
    records.Add "LRV|8.4 Mitigation & Control|D338|A1203A|A12031|EnumerationFrequency|B338|17478703"
    records.Add "LRW|8.4 Mitigation & Control|C340|A1203A|A12031|EnumerationYesNo|B340|302693416"
    records.Add "LRX|8.4 Mitigation & Control|C341|A1203A|A12031|EnumerationYesNo|B341|464289613"
    records.Add "LRY|8.4 Mitigation & Control|C342|A1203A|A12031|EnumerationFrequency|B342|657263045"
    records.Add "LRZ|8.4 Mitigation & Control|C343|A1203A|A12031|EnumerationYesNo|B343|1330788433"
    records.Add "LSA|8.4 Mitigation & Control|C344|A1203A|A12031|EnumerationYesNoNA|B344|1387572890"
    records.Add "LSB|8.4 Mitigation & Control|C345|A1203A|A12031|String500|B345|123317415"
    records.Add "LSC|8.4 Mitigation & Control|C351|A1203A|A12031|PercentageValue|B351|66563"
    records.Add "LSD|8.4 Mitigation & Control|C352|A1203A|A12031|PercentageValue|B352|158731779"
    records.Add "LSE|8.4 Mitigation & Control|C353|A1203A|A12031|PercentageValue|B353|2556034"
    records.Add "LSF|8.4 Mitigation & Control|C354|A1203A|A12031|PercentageValue|B354|1361244929"
    records.Add "LSG|8.4 Mitigation & Control|C355|A1203A|A12031|PercentageValue|B355|841775554"
    records.Add "LSH|8.4 Mitigation & Control|C357|A1203A|A12031|PercentageValue|B357|2119956532"
    records.Add "LSI|8.4 Mitigation & Control|C359|A1203A|A12031|IntegerValue|B359|375912221"
    records.Add "LSJ|8.4 Mitigation & Control|C360|A1203A|A12031|IntegerValue|B360|264339013"
End Sub

Private Sub LoadTableMap(ByRef records As Collection)
    records.Add "A1201B|8.2 General|28|29|58|B,D,E,F,G|LAR,LAS,LAT,LAU,LAV"
    records.Add "A1202A|8.3.1 Inherent Risk Funds|4|5|0|A,B,C,D,E,F,G,H,I,J,K,L,M,N,O|LBB,LBC,LBD,LBE,LBF,LBG,LBH,LBI,LBJ,LBK,LBL,LBM,LBN,LBO,LBP"
    records.Add "A1202B|8.3.2 Inherent Risks Investors|7|8|34|B,D|LBQ,LBR"
    records.Add "A1204|8.5 Physical Presence|3|4|0|A,B,C,D,E|LSK,LSL,LSM,LSN,LSO"
    records.Add "A1205|8.6 Residence & Establishment|4|5|0|A,B,C,D,E,F,G,H,I|LSP,LSQ,LSR,LSS,LST,LSU,LSV,LSW,LSX"
    records.Add "A1206|8.7 Beneficial Owner|4|5|0|A,B|LSY,LSZ"
    records.Add "A1207|8.8 Politically Exposed Person|3|4|0|A,B,C|LTA,LTB,LTC"
    records.Add "A1208|8.9 Geography of Funds Flow|4|5|0|A,B,C,D,E,F,G,H,I,J,K,L,M|LTD,LTE,LTF,LTG,LTH,LTI,LTJ,LTK,LTL,LTM,LTN,LTO,LTP"
    records.Add "A1209|8.10 Transaction Monitoring|3|4|0|A,B,C,D,E|LTQ,LTR,LTS,LTT,LTU"
End Sub
