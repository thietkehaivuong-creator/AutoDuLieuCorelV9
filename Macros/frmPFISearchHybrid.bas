Option Explicit

' ==============================================================
' AUTO DU LIEU COREL V9 - HYBRID MODE (API + LOCAL TSV)
' Macro: frmPFISearchHybrid
' Reads: PowerShell API data (TSV/JSON) + Local demo data
' ==============================================================

Public Const DATA_FOLDER As String = "AutoDuLieuCorelV9"
Public Const API_TSV_FILE As String = "auto_du_lieu_corel_api_v9.tsv"
Public Const DEMO_TSV_FILE As String = "auto_du_lieu_corel_demo_v9.tsv"
Public Const FAO_MAP_FILE As String = "fao_mapping_v9.tsv"
Public Const PPDB_MAP_FILE As String = "fishing_method_mapping_v9.tsv"

Private WithEvents frmMain As Object
Private lstCustomers As Object
Private lstPFI As Object
Private lstProducts As Object
Private txtSearch As Object
Private txtDetail As Object
Private btnSearch As Object
Private btnApply As Object
Private btnSync As Object
Private btnSetting As Object
Private btnClose As Object
Private cmbMode As Object
Private cmbDateFormat As Object
Private lblStatus As Object
Private lblDataMode As Object
Private chkUseAPI As Object

Private currentRows As Collection
Private selectedRow As Object
Private allRows As Collection
Private usingAPIData As Boolean

Public Sub AutoDuLieuCorelV9_Hybrid_Show()
    BuildForm
    frmMain.Show vbModeless
End Sub

Private Function AppFolder() As String
    Dim p As String
    p = Environ$("APPDATA") & "\" & DATA_FOLDER
    If Dir(p, vbDirectory) = "" Then MkDir p
    AppFolder = p
End Function

Private Function FilePath(ByVal fileName As String) As String
    FilePath = AppFolder() & "\" & fileName
End Function

Private Sub BuildForm()
    Dim frm As Object
    Set frm = CreateObject("Scripting.Dictionary")
    
    With CreateObject("ADODB.Stream")
        .Charset = "utf-8"
        .Open
        .WriteText CreateFormHTML()
        .SaveToFile FilePath("form.html"), 2
        .Close
    End With
    
    Set frmMain = CreateObject("Shell.BrowserWindow")
    Call frmMain.Navigate(FilePath("form.html"))
End Sub

Private Function CreateFormHTML() As String
    Dim html As String
    html = "<!DOCTYPE html>" & vbCrLf & _
           "<html><head><meta charset='utf-8'>" & vbCrLf & _
           "<title>AUTO DU LIEU COREL V9</title>" & vbCrLf & _
           "<style>" & vbCrLf & _
           "body { background: #262626; color: #f5f5f5; font-family: Segoe UI; margin: 0; padding: 10px; }" & vbCrLf & _
           ".container { max-width: 1200px; margin: 0 auto; }" & vbCrLf & _
           ".header { background: #1a1a1a; padding: 15px; border-radius: 4px; margin-bottom: 10px; }" & vbCrLf & _
           ".header h1 { margin: 0; font-size: 18px; }" & vbCrLf & _
           ".status { font-size: 12px; color: #a0dca0; margin-top: 5px; }" & vbCrLf & _
           ".controls { background: #1a1a1a; padding: 10px; margin-bottom: 10px; border-radius: 4px; }" & vbCrLf & _
           ".controls input, .controls select, .controls button { padding: 8px; margin-right: 5px; background: #1e1e1e; color: #f5f5f5; border: 1px solid #737373; border-radius: 3px; font-size: 12px; }" & vbCrLf & _
           ".controls button { background: #444; cursor: pointer; }" & vbCrLf & _
           ".controls button:hover { background: #555; }" & vbCrLf & _
           ".grid { display: grid; grid-template-columns: 1fr 1fr 1fr 1.5fr; gap: 10px; margin-bottom: 10px; }" & vbCrLf & _
           ".panel { background: #1a1a1a; border: 1px solid #444; border-radius: 4px; padding: 10px; }" & vbCrLf & _
           ".panel h3 { margin: 0 0 10px 0; font-size: 13px; color: #a0dca0; }" & vbCrLf & _
           ".panel ul { list-style: none; margin: 0; padding: 0; height: 200px; overflow-y: auto; }" & vbCrLf & _
           ".panel li { padding: 5px; cursor: pointer; border-radius: 2px; font-size: 12px; }" & vbCrLf & _
           ".panel li:hover { background: #333; }" & vbCrLf & _
           ".panel li.selected { background: #0078d4; }" & vbCrLf & _
           ".preview { grid-column: 1/-1; background: #1a1a1a; border: 1px solid #444; border-radius: 4px; padding: 10px; }" & vbCrLf & _
           ".preview textarea { width: 100%; height: 200px; background: #0a0a0a; color: #a0dca0; border: 1px solid #444; padding: 8px; font-family: Consolas; font-size: 11px; border-radius: 3px; resize: none; }" & vbCrLf & _
           ".footer { background: #1a1a1a; padding: 10px; border-radius: 4px; display: flex; gap: 5px; }" & vbCrLf & _
           ".footer button { flex: 1; padding: 10px; }" & vbCrLf & _
           "</style></head><body>" & vbCrLf & _
           "<div class='container'>" & vbCrLf & _
           "<div class='header'>" & vbCrLf & _
           "<h1>AUTO DU LIEU COREL V9 - HYBRID MODE</h1>" & vbCrLf & _
           "<div class='status' id='status'>Dang tai du lieu...</div>" & vbCrLf & _
           "</div>" & vbCrLf & _
           "<div class='controls'>" & vbCrLf & _
           "<input type='text' id='search' placeholder='Tim PFI / Khach / San pham / Lot' style='width: 300px;'>" & vbCrLf & _
           "<button onclick='doSearch()'>TIM</button>" & vbCrLf & _
           "<select id='dateFormat'>" & vbCrLf & _
           "<option value='dd/mm/yyyy'>dd/mm/yyyy</option>" & vbCrLf & _
           "<option value='mm/dd/yyyy'>mm/dd/yyyy</option>" & vbCrLf & _
           "<option value='yyyy/mm/dd'>yyyy/mm/dd</option>" & vbCrLf & _
           "<option value='mm/yyyy'>mm/yyyy</option>" & vbCrLf & _
           "<option value='yyyy-mm-dd'>yyyy-mm-dd</option>" & vbCrLf & _
           "</select>" & vbCrLf & _
           "<button onclick='syncData()'>NAP DU LIEU</button>" & vbCrLf & _
           "<button onclick='openSettings()'>CAI DAT</button>" & vbCrLf & _
           "<button onclick='closeForm()'>DONG</button>" & vbCrLf & _
           "</div>" & vbCrLf & _
           "<div class='grid'>" & vbCrLf & _
           "<div class='panel'><h3>1. KHACH HANG</h3><ul id='customers'></ul></div>" & vbCrLf & _
           "<div class='panel'><h3>2. PFI</h3><ul id='pfis'></ul></div>" & vbCrLf & _
           "<div class='panel'><h3>3. SAN PHAM</h3><ul id='products'></ul></div>" & vbCrLf & _
           "<div class='panel'><h3>4. LOT / THONG TIN</h3><ul id='details'></ul></div>" & vbCrLf & _
           "<div class='preview'><h3>CHI TIET DU LIEU</h3><textarea id='preview' readonly></textarea></div>" & vbCrLf & _
           "</div>" & vbCrLf & _
           "<div class='footer'>" & vbCrLf & _
           "<button onclick='applyToCorel()' style='background: #0078d4;'>DO DU LIEU VAO TEM COREL</button>" & vbCrLf & _
           "</div>" & vbCrLf & _
           "</div>" & vbCrLf & _
           "<script>" & vbCrLf & _
           "let allData = [];" & vbCrLf & _
           "let selectedCustomer = null;" & vbCrLf & _
           "let selectedPFI = null;" & vbCrLf & _
           "let selectedProduct = null;" & vbCrLf & _
           "function initForm() { loadData(); }" & vbCrLf & _
           "function loadData() { }" & vbCrLf & _
           "function doSearch() { }" & vbCrLf & _
           "function syncData() { }" & vbCrLf & _
           "function applyToCorel() { }" & vbCrLf & _
           "function openSettings() { }" & vbCrLf & _
           "function closeForm() { window.close(); }" & vbCrLf & _
           "window.onload = initForm;" & vbCrLf & _
           "</script>" & vbCrLf & _
           "</body></html>"
    CreateFormHTML = html
End Function

' Simplified VBA Form Builder
Private Sub BuildForm_VBA()
    Dim frm As Object
    Set frm = CreateObject("ADODB.Recordset")
    
    ' Create form object manually
    Dim newForm As New UserForm1
    newForm.Show vbModeless
End Sub

Public Sub TestLoadData()
    Set allRows = LoadTSVData(FilePath(API_TSV_FILE))
    If allRows.Count = 0 Then
        Set allRows = LoadTSVData(FilePath(DEMO_TSV_FILE))
        usingAPIData = False
        MsgBox "Dang su dung DEMO data (khong tim thay API data)", vbInformation
    Else
        usingAPIData = True
        MsgBox "Tim thay API data: " & allRows.Count & " dong", vbInformation
    End If
    
    RefreshCustomerList
End Sub

Private Function LoadTSVData(ByVal filePath As String) As Collection
    Dim rows As New Collection
    Dim f As Integer, line As String, arr As Variant, headers As Variant
    Dim d As Object, i As Long, first As Boolean
    
    On Error GoTo bad
    
    If Dir(filePath) = "" Then
        Set LoadTSVData = rows
        Exit Function
    End If
    
    f = FreeFile
    Open filePath For Input As #f
    first = True
    
    Do Until EOF(f)
        Line Input #f, line
        line = Trim(line)
        If line <> "" Then
            arr = Split(line, vbTab)
            If first Then
                headers = arr
                first = False
            Else
                Set d = CreateObject("Scripting.Dictionary")
                For i = LBound(headers) To UBound(headers)
                    If i <= UBound(arr) Then
                        d(CStr(headers(i))) = CStr(arr(i))
                    Else
                        d(CStr(headers(i))) = ""
                    End If
                Next i
                rows.Add d
            End If
        End If
    Loop
    
    Close #f
    Set LoadTSVData = rows
    Exit Function
    
bad:
    On Error Resume Next
    Close #f
    Set LoadTSVData = New Collection
End Function

Private Sub RefreshCustomerList()
    Dim customers As New Collection, seen As Object, i As Long, cust As String
    Set seen = CreateObject("Scripting.Dictionary")
    
    If allRows Is Nothing Then Exit Sub
    
    For i = 1 To allRows.Count
        cust = SafeGet(allRows(i), "Customer")
        If Not seen.Exists(UCase(cust)) Then
            seen.Add UCase(cust), True
            customers.Add cust
        End If
    Next i
    
    ' Update UI with customer list
End Sub

Private Function SafeGet(ByVal d As Object, ByVal key As String) As String
    On Error Resume Next
    If d.Exists(key) Then
        SafeGet = CStr(d(key))
    Else
        SafeGet = ""
    End If
End Function

Private Function SearchData(ByVal q As String) As Collection
    Dim results As New Collection, i As Long, hay As String
    
    If allRows Is Nothing Then Set SearchData = results: Exit Function
    
    q = UCase(Trim(q))
    
    For i = 1 To allRows.Count
        hay = UCase(SafeGet(allRows(i), "Customer") & " " & _
                    SafeGet(allRows(i), "PFI") & " " & _
                    SafeGet(allRows(i), "Description") & " " & _
                    SafeGet(allRows(i), "Lot_No"))
        
        If q = "" Or InStr(1, hay, q, vbTextCompare) > 0 Then
            results.Add allRows(i)
        End If
    Next i
    
    Set SearchData = results
End Function

Private Sub InsertToCorel()
    On Error GoTo bad
    
    If selectedRow Is Nothing Then
        MsgBox "Hay chon san pham truoc", vbExclamation
        Exit Sub
    End If
    
    Dim text As String
    text = "CUSTOMER: " & SafeGet(selectedRow, "Customer") & vbCrLf & _
           "PFI: " & SafeGet(selectedRow, "PFI") & vbCrLf & _
           "PRODUCT: " & SafeGet(selectedRow, "Description") & vbCrLf & _
           "SIZE: " & SafeGet(selectedRow, "Sizes") & vbCrLf & _
           "QTY: " & SafeGet(selectedRow, "Qty") & vbCrLf & _
           "LOT: " & SafeGet(selectedRow, "Lot_No") & vbCrLf & _
           "FAO: " & SafeGet(selectedRow, "FAO") & vbCrLf & _
           "PACKING: " & SafeGet(selectedRow, "Packing")
    
    ActiveLayer.CreateArtisticText 0, 0, text
    MsgBox "Da insert du lieu vao Corel!", vbInformation
    Exit Sub
    
bad:
    MsgBox "Loi: " & Err.Description, vbExclamation
End Sub

' Utility Functions
Private Function ParseDateFlexible(ByVal s As String) As Date
    Dim arr As Variant, d As Long, m As Long, Y As Long
    s = Trim(Replace(Replace(s, "-", "/"), ".", "/"))
    
    If InStr(s, "/") > 0 Then
        arr = Split(s, "/")
        If UBound(arr) = 1 Then
            m = CLng(Val(arr(0))): Y = CLng(Val(arr(1)))
            If Y < 100 Then Y = 2000 + Y
            ParseDateFlexible = DateSerial(Y, m, 1)
        ElseIf UBound(arr) >= 2 Then
            If Len(Trim(CStr(arr(0)))) = 4 Then
                Y = CLng(Val(arr(0))): m = CLng(Val(arr(1))): d = CLng(Val(arr(2)))
            Else
                d = CLng(Val(arr(0))): m = CLng(Val(arr(1))): Y = CLng(Val(arr(2)))
            End If
            If Y < 100 Then Y = 2000 + Y
            ParseDateFlexible = DateSerial(Y, m, d)
        End If
    Else
        ParseDateFlexible = CDate(s)
    End If
End Function

Private Function FormatDate(ByVal s As String, ByVal fmt As String) As String
    On Error GoTo bad
    Dim dt As Date
    If Trim(s) = "" Or Trim(fmt) = "" Then FormatDate = s: Exit Function
    
    dt = ParseDateFlexible(s)
    FormatDate = Format$(dt, fmt)
    Exit Function
    
bad:
    FormatDate = s
End Function
