Option Explicit

' ============================================================
' AUTO DU LIEU COREL V9 - PFI SEARCH MACRO
' Giao diện VBA cho CorelDRAW
' ============================================================

Private apiBaseUrl As String
Private currentData As Object
Private currentDateFormats As Collection

Private Sub UserForm_Initialize()
    apiBaseUrl = "http://localhost:5000/api"
    
    ' Form Setup
    Me.Caption = "AUTO DU LIEU - PFI SEARCH"
    Me.BackColor = RGB(38, 38, 38)
    Me.ForeColor = RGB(240, 240, 240)
    Me.Width = 900
    Me.Height = 700
    
    ' ====== TITLE ======
    Dim lblTitle As MSForms.Label
    Set lblTitle = Me.Controls.Add("Forms.Label.1", "lblTitle")
    lblTitle.Caption = "AUTO DU LIEU - TIM KIEM PFI"
    lblTitle.Font.Bold = True
    lblTitle.Font.Size = 14
    lblTitle.BackColor = RGB(38, 38, 38)
    lblTitle.ForeColor = RGB(255, 255, 255)
    lblTitle.Move 10, 10, 350, 28
    
    ' ====== SEARCH SECTION ======
    Dim lblPFI As MSForms.Label
    Set lblPFI = Me.Controls.Add("Forms.Label.1", "lblPFI")
    lblPFI.Caption = "Nhap PFI (Don hang):"
    lblPFI.BackColor = RGB(38, 38, 38)
    lblPFI.ForeColor = RGB(240, 240, 240)
    lblPFI.Move 10, 50, 150, 20
    
    Dim txtPFI As MSForms.TextBox
    Set txtPFI = Me.Controls.Add("Forms.TextBox.1", "txtPFI")
    txtPFI.BackColor = RGB(30, 30, 30)
    txtPFI.ForeColor = RGB(245, 245, 245)
    txtPFI.Move 160, 48, 200, 24
    
    Dim btnSearch As MSForms.CommandButton
    Set btnSearch = Me.Controls.Add("Forms.CommandButton.1", "btnSearch")
    btnSearch.Caption = "TIM KIEM"
    btnSearch.BackColor = RGB(68, 68, 68)
    btnSearch.ForeColor = RGB(245, 245, 245)
    btnSearch.Move 370, 48, 80, 24
    
    Dim lblStatus As MSForms.Label
    Set lblStatus = Me.Controls.Add("Forms.Label.1", "lblStatus")
    lblStatus.Caption = "San sang"
    lblStatus.BackColor = RGB(38, 38, 38)
    lblStatus.ForeColor = RGB(160, 220, 160)
    lblStatus.Move 460, 52, 250, 18
    
    ' ====== DATA GRID ======
    Dim lblDataTitle As MSForms.Label
    Set lblDataTitle = Me.Controls.Add("Forms.Label.1", "lblDataTitle")
    lblDataTitle.Caption = "Du lieu san pham:"
    lblDataTitle.Font.Bold = True
    lblDataTitle.BackColor = RGB(38, 38, 38)
    lblDataTitle.ForeColor = RGB(240, 240, 240)
    lblDataTitle.Move 10, 85, 200, 20
    
    Dim gridData As MSForms.ListBox
    Set gridData = Me.Controls.Add("Forms.ListBox.1", "gridData")
    gridData.ColumnCount = 2
    gridData.ColumnWidths = "200;400"
    gridData.Move 10, 110, 720, 280
    gridData.BackColor = RGB(30, 30, 30)
    gridData.ForeColor = RGB(245, 245, 245)
    
    ' ====== FORMAT SECTION ======
    Dim lblFormatTitle As MSForms.Label
    Set lblFormatTitle = Me.Controls.Add("Forms.Label.1", "lblFormatTitle")
    lblFormatTitle.Caption = "Dinh dang ngay (cho cac cot ngay):"
    lblFormatTitle.Font.Bold = True
    lblFormatTitle.BackColor = RGB(38, 38, 38)
    lblFormatTitle.ForeColor = RGB(240, 240, 240)
    lblFormatTitle.Move 10, 400, 300, 20
    
    Dim cmbDateFormat As MSForms.ComboBox
    Set cmbDateFormat = Me.Controls.Add("Forms.ComboBox.1", "cmbDateFormat")
    cmbDateFormat.Move 10, 425, 180, 24
    cmbDateFormat.BackColor = RGB(30, 30, 30)
    cmbDateFormat.ForeColor = RGB(245, 245, 245)
    PopulateDateFormats cmbDateFormat
    cmbDateFormat.Text = "dd/mm/yyyy"
    
    Dim lblCustomHint As MSForms.Label
    Set lblCustomHint = Me.Controls.Add("Forms.Label.1", "lblCustomHint")
    lblCustomHint.Caption = "Custom: mm/yyyy, yyyy-mm-dd, MMM.yyyy"
    lblCustomHint.BackColor = RGB(38, 38, 38)
    lblCustomHint.ForeColor = RGB(200, 200, 200)
    lblCustomHint.Font.Size = 8
    lblCustomHint.Move 200, 428, 300, 18
    
    ' ====== ACTION BUTTONS ======
    Dim btnCopyAll As MSForms.CommandButton
    Set btnCopyAll = Me.Controls.Add("Forms.CommandButton.1", "btnCopyAll")
    btnCopyAll.Caption = "COPY TAT CA"
    btnCopyAll.BackColor = RGB(68, 68, 68)
    btnCopyAll.ForeColor = RGB(245, 245, 245)
    btnCopyAll.Move 10, 460, 100, 26
    
    Dim btnCopySelected As MSForms.CommandButton
    Set btnCopySelected = Me.Controls.Add("Forms.CommandButton.1", "btnCopySelected")
    btnCopySelected.Caption = "COPY DONG"
    btnCopySelected.BackColor = RGB(68, 68, 68)
    btnCopySelected.ForeColor = RGB(245, 245, 245)
    btnCopySelected.Move 120, 460, 100, 26
    
    Dim btnInsertCorel As MSForms.CommandButton
    Set btnInsertCorel = Me.Controls.Add("Forms.CommandButton.1", "btnInsertCorel")
    btnInsertCorel.Caption = "INSERT COREL"
    btnInsertCorel.BackColor = RGB(68, 68, 68)
    btnInsertCorel.ForeColor = RGB(245, 245, 245)
    btnInsertCorel.Move 230, 460, 100, 26
    
    Dim btnSettings As MSForms.CommandButton
    Set btnSettings = Me.Controls.Add("Forms.CommandButton.1", "btnSettings")
    btnSettings.Caption = "CAI DAT"
    btnSettings.BackColor = RGB(68, 68, 68)
    btnSettings.ForeColor = RGB(245, 245, 245)
    btnSettings.Move 340, 460, 80, 26
    
    Dim btnClose As MSForms.CommandButton
    Set btnClose = Me.Controls.Add("Forms.CommandButton.1", "btnClose")
    btnClose.Caption = "DONG"
    btnClose.BackColor = RGB(68, 68, 68)
    btnClose.ForeColor = RGB(245, 245, 245)
    btnClose.Move 430, 460, 80, 26
End Sub

Private Sub PopulateDateFormats(cmbFormat As MSForms.ComboBox)
    Set currentDateFormats = New Collection
    
    Dim formats As Variant
    formats = Array( _
        "dd/mm/yyyy", _
        "d/m/yyyy", _
        "mm/dd/yyyy", _
        "m/d/yyyy", _
        "yyyy/mm/dd", _
        "yyyy/m/d", _
        "dd-mm-yyyy", _
        "mm-dd-yyyy", _
        "yyyy-mm-dd", _
        "dd.mm.yyyy", _
        "yyyy.mm.dd", _
        "mm/yyyy", _
        "m/yyyy", _
        "yyyy/mm", _
        "yyyy-mm", _
        "mmm yyyy", _
        "mmmm yyyy", _
        "dd mmm yyyy", _
        "mmm dd, yyyy" _
    )
    
    Dim i As Long
    For i = LBound(formats) To UBound(formats)
        cmbFormat.AddItem CStr(formats(i))
        currentDateFormats.Add CStr(formats(i))
    Next i
End Sub

Private Sub btnSearch_Click()
    Dim pfi As String
    pfi = Trim(Me.Controls("txtPFI").Text)
    
    If pfi = "" Then
        MsgBox "Hay nhap PFI", vbExclamation
        Exit Sub
    End If
    
    Me.Controls("lblStatus").Caption = "Dang tim kiem..."
    Application.ScreenUpdating = False
    
    On Error GoTo errHandler
    
    Dim url As String
    url = apiBaseUrl & "/data/search/" & URLEncode(pfi)
    
    Dim response As String
    response = CallAPI("GET", url, "")
    
    If InStr(response, """success"":true") > 0 Then
        DisplayData response, pfi
        Me.Controls("lblStatus").Caption = "Tim thay du lieu cho PFI: " & pfi
    Else
        MsgBox "Khong tim thay PFI: " & pfi, vbExclamation
        Me.Controls("lblStatus").Caption = "Khong tim thay: " & pfi
        Me.Controls("gridData").Clear
        Set currentData = Nothing
    End If
    
    Application.ScreenUpdating = True
    Exit Sub
    
errHandler:
    MsgBox "Loi tim kiem: " & Err.Description, vbCritical
    Me.Controls("lblStatus").Caption = "Loi"
End Sub

Private Sub DisplayData(jsonResponse As String, pfi As String)
    Dim gridData As MSForms.ListBox
    Set gridData = Me.Controls("gridData")
    gridData.Clear
    
    ' Parse JSON thong tin (simplified)
    Dim headers As Variant
    headers = Array("Customer", "PFI", "Lo", "Description", "Sizes", "Qty", "SPEC", _
                    "Xn_Tem_Hang", "Ng_DG_tu", "Ng_DG", "Best_Before", "Nguon_NL", _
                    "FAO", "haTGDB", "PP_DB", "Lot_No", "Remark", "NLT_ve", "TGDB_NK", "Packing")
    
    ' Hien thi du lieu tu demo
    Dim i As Long
    Dim fieldValue As String
    
    For i = LBound(headers) To UBound(headers)
        Dim fieldName As String
        fieldName = CStr(headers(i))
        
        ' Gia su data duoc tra ve tu API
        fieldValue = ExtractJSONField(jsonResponse, fieldName)
        
        ' Format dates
        If IsDateField(fieldName) And fieldValue <> "" Then
            Dim targetFormat As String
            targetFormat = Me.Controls("cmbDateFormat").Text
            fieldValue = FormatDateViaAPI(fieldValue, targetFormat)
        End If
        
        gridData.AddItem fieldName
        gridData.List(gridData.ListCount - 1, 1) = fieldValue
    Next i
End Sub

Private Function ExtractJSONField(json As String, fieldName As String) As String
    ' Simplified JSON extraction - find "fieldname":"value"
    Dim pattern As String
    pattern = """" & fieldName & """:""([^""]*)"""
    
    Dim start As Long
    start = InStr(json, pattern)
    
    If start = 0 Then
        ExtractJSONField = ""
        Exit Function
    End If
    
    Dim beginQuote As Long
    beginQuote = InStr(start, json, """") + 1
    
    Dim endQuote As Long
    endQuote = InStr(beginQuote, json, """")
    
    If endQuote > beginQuote Then
        ExtractJSONField = Mid(json, beginQuote, endQuote - beginQuote)
    Else
        ExtractJSONField = ""
    End If
End Function

Private Function IsDateField(fieldName As String) As Boolean
    Select Case UCase(fieldName)
        Case "XN_TEM_HANG", "NG_DG_TU", "NG_DG", "BEST_BEFORE", "HATGDB", "NLT_VE", "TGDB_NK"
            IsDateField = True
        Case Else
            IsDateField = False
    End Select
End Function

Private Function FormatDateViaAPI(inputDate As String, targetFormat As String) As String
    On Error GoTo errHandler
    
    Dim url As String
    url = apiBaseUrl & "/data/format-date"
    
    Dim payload As String
    payload = "{""inputDate"":""" & inputDate & """,""targetFormat"":""" & targetFormat & """}"
    
    Dim response As String
    response = CallAPI("POST", url, payload)
    
    If InStr(response, """success"":true") > 0 Then
        ' Extract formatted date from response
        Dim formattedDate As String
        formattedDate = ExtractJSONField(response, "formattedDate")
        If formattedDate <> "" Then
            FormatDateViaAPI = formattedDate
        Else
            FormatDateViaAPI = inputDate
        End If
    Else
        FormatDateViaAPI = inputDate
    End If
    
    Exit Function
errHandler:
    FormatDateViaAPI = inputDate
End Function

Private Sub btnCopyAll_Click()
    Dim gridData As MSForms.ListBox
    Set gridData = Me.Controls("gridData")
    
    If gridData.ListCount = 0 Then
        MsgBox "Khong co du lieu de copy", vbExclamation
        Exit Sub
    End If
    
    Dim output As String
    output = ""
    
    Dim i As Long
    For i = 0 To gridData.ListCount - 1
        Dim fieldName As String
        Dim fieldValue As String
        fieldName = gridData.List(i, 0)
        fieldValue = gridData.List(i, 1)
        
        output = output & fieldName & ": " & fieldValue & vbCrLf
    Next i
    
    CopyToClipboard output
    MsgBox "Da copy tat ca du lieu (" & gridData.ListCount & " dong)", vbInformation
End Sub

Private Sub btnCopySelected_Click()
    Dim gridData As MSForms.ListBox
    Set gridData = Me.Controls("gridData")
    
    If gridData.ListIndex < 0 Then
        MsgBox "Chon mot dong truoc", vbExclamation
        Exit Sub
    End If
    
    Dim fieldName As String
    Dim fieldValue As String
    
    fieldName = gridData.List(gridData.ListIndex, 0)
    fieldValue = gridData.List(gridData.ListIndex, 1)
    
    CopyToClipboard fieldValue
    MsgBox "Da copy: " & fieldName & " = " & fieldValue, vbInformation
End Sub

Private Sub btnInsertCorel_Click()
    Dim gridData As MSForms.ListBox
    Set gridData = Me.Controls("gridData")
    
    If gridData.ListCount = 0 Then
        MsgBox "Khong co du lieu", vbExclamation
        Exit Sub
    End If
    
    On Error GoTo errHandler
    
    Dim doc As Document
    Set doc = ActiveDocument
    
    If doc Is Nothing Then
        MsgBox "Khong co document CorelDRAW nao mo", vbExclamation
        Exit Sub
    End If
    
    ' Insert text vao CorelDRAW
    Dim pfi As String
    Dim customer As String
    Dim description As String
    Dim text As String
    
    pfi = gridData.List(1, 1) ' PFI la dong 2
    customer = gridData.List(0, 1) ' Customer la dong 1
    description = gridData.List(3, 1) ' Description la dong 4
    
    text = "PFI: " & pfi & vbCrLf & _
           "Customer: " & customer & vbCrLf & _
           "Product: " & description
    
    Dim layer As Layer
    Set layer = doc.ActiveLayer
    
    Dim shape As Shape
    Set shape = layer.CreateArtisticText(0, 0, text)
    
    MsgBox "Da insert du lieu vao CorelDRAW", vbInformation
    
    Exit Sub
errHandler:
    MsgBox "Loi insert: " & Err.Description, vbCritical
End Sub

Private Sub btnSettings_Click()
    MsgBox "CAI DAT THONG TIN DANG NHAP:" & vbCrLf & vbCrLf & _
           "1. Mo trình duyệt: http://localhost:5000/swagger" & vbCrLf & _
           "2. Click: POST /api/settings/credentials" & vbCrLf & _
           "3. Nhap username va password" & vbCrLf & _
           "4. Click 'Execute'", vbInformation, "CAI DAT"
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

' ============================================================
' UTILITY FUNCTIONS
' ============================================================

Private Function CallAPI(method As String, url As String, payload As String) As String
    Dim xmlHttp As Object
    Set xmlHttp = CreateObject("MSXML2.XMLHTTP")
    
    On Error GoTo errHandler
    
    xmlHttp.Open method, url, False
    xmlHttp.SetRequestHeader "Content-Type", "application/json"
    xmlHttp.SetRequestHeader "Accept", "application/json"
    
    If method = "POST" Then
        xmlHttp.Send payload
    Else
        xmlHttp.Send
    End If
    
    CallAPI = xmlHttp.responseText
    
    Exit Function
errHandler:
    MsgBox "HTTP Error: " & Err.Description, vbCritical
    CallAPI = ""
End Function

Private Sub CopyToClipboard(text As String)
    On Error Resume Next
    With CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        .SetText text
        .PutInClipboard
    End With
End Sub

Private Function URLEncode(text As String) As String
    URLEncode = Replace(text, " ", "%20")
    URLEncode = Replace(URLEncode, "&", "%26")
    URLEncode = Replace(URLEncode, "=", "%3D")
End Function
