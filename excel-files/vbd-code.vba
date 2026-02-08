Sub SendAllPayslip()
Dim Data As Worksheet, PS As Worksheet, Body As Worksheet
Dim lr As Long
ActiveSheet.Outline.ShowLevels RowLevels:=0, ColumnLevels:=1
Range("A4").Select


'Nhap thang gui TBKQ
my_date = DateSerial(2016, 3, 31)
      
ShowInputBox1:
thang_gui = Application.InputBox("Vui long nhap THANG gui TBKQ (mm/yyyy): " & vbCrLf & " ", Title:="Nhap THANG cua TBKQ")

If thang_gui = False Then
    Exit Sub
End If

thang_thang = Left(thang_gui, 2)
nam_nam = Right(thang_gui, 4)

If Len(thang_gui) <> 7 Or thang_thang > 12 Or thang_thang <= 0 Or Mid(thang_gui, 3, 1) <> "/" Then
    MsgBox ("THANG ban nhap: '" & thang_gui & "', khong dung format 'mm/yyyy', vui long nhap lai.")
    GoTo ShowInputBox1
End If
            
yr_date = DateSerial(nam_nam, thang_thang, 1)
sys_date = DateSerial(Year(Now()), Month(Now()), Day(Now()))

    
'Tao duong dan
duongdan = ActiveWorkbook.Path & "\"

If Right(duongdan, 8) <> "Payslip\" Then
    MsgBox ("Chua tao thu muc Payslip de luu file gui phieu luong !")
    Exit Sub
End If

'Kiem tra password: chua co password
Sheets("Data").Select
    lr = Range("A" & Rows.Count).End(xlUp).Row

Set rFound = Cells.Find(What:="PassWord", after:=ActiveCell, LookIn:=xlFormulas, _
    LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=True, SearchFormat:=False)
      
If rFound Is Nothing Then
    MsgBox ("Khong tim thay cot 'PassWord' , Chuong trinh se thoat ra.")
    Range("A2").Select
    Exit Sub
End If

Cells.Find(What:="PassWord", after:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=True, SearchFormat:=False).Activate
    
colpw = Split(ActiveCell.Address, "$")(1)
cellpw = ActiveCell.Address
    
If WorksheetFunction.CountBlank(Range(colpw & "1:" & colpw & lr)) > 0 Then
    Range(cellpw).Select
    MsgBox ("Co NV chua co password cho phieu luong, Chuong trinh se thoat ra de ban kiem tra lai.")
    Exit Sub
End If
    
'Kiem tra danh sach email: chua co email, email giong nhau
Sheets("Data").Select
Set rFound = Cells.Find(What:="EmailAddress", after:=ActiveCell, LookIn:=xlFormulas, _
    LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=True, SearchFormat:=False)
      
If rFound Is Nothing Then
    MsgBox ("Khong tim thay cot 'EmailAddress' , Chuong trinh se thoat ra.")
    Range("A2").Select
    Exit Sub
End If

Cells.Find(What:="EmailAddress", after:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=True, SearchFormat:=False).Activate
    
colemail = Split(ActiveCell.Address, "$")(1)
cellemail = ActiveCell.Address
    
If WorksheetFunction.CountBlank(Range(colemail & "4:" & colemail & lr)) > 0 Then
    Range(cellemail).Select
    MsgBox ("Co NV chua co email, Chuong trinh se thoat ra de ban kiem tra lai.")
    Exit Sub
End If
    
For i = 4 To lr
    kt1 = i
    kt2 = kt1 + 1
        
    If WorksheetFunction.CountIf(Range(colemail & "4:" & colemail & lr), Range(colemail & kt1)) > 1 Then
        Range(colemail & kt1).Select
        MsgBox ("Cot 'EMAIL': o " & colemail & kt1 & " co dia chi email giong cua NV khac, Chuong trinh se thoat ra.")
        Exit Sub
    End If
Next i

'Tao Name 'TBKQ'
Cells.Find(What:="MNV", after:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=True, SearchFormat:=False).Activate
        
colTBKQ = Split(ActiveCell.Address, "$")(1)
ActiveWorkbook.Names.Add Name:="TBKQ", RefersTo:="=Data!$A$4:$" & colTBKQ & "$" & lr

'Xoa nhung Name rac
Dim NamedRange As Name
For Each NamedRange In ActiveWorkbook.Names
    If NamedRange.Name <> "TBKQ" Then
        NamedRange.Delete
    End If
Next NamedRange


tt_TBKQ = i

'Thong bao bat dau gui emai
If MsgBox("Chuong trinh da hoan tat viec kiem tra du lieu." _
    & vbNewLine & " o Thang gui Phieu luong     : " & thang_gui _
    & vbNewLine & " o Tong so Phieu luong       : " & tt_payslip - 4 _
    & vbNewLine & "Ban co chac la da mo MS Outlook va muon gui Phieu luong ngay bay gio ?", vbYesNo + vbDefaultButton2) = vbNo Then
    Exit Sub
End If

Set Data = Worksheets("Data")
Set PS = Worksheets("TBKQ")
Set Body = Worksheets("bodymail")

'Tao Subject và Body message
sub1 = PS.Range("G1").Value
body1 = Body.Range("A1").Value
body3 = Body.Range("A3").Value
body4 = Body.Range("A4").Value
body5 = Body.Range("A5").Value
body7 = Body.Range("A7").Value
body8 = Body.Range("A8").Value
body9 = Body.Range("A9").Value
body11 = Body.Range("A11").Value
body12 = Body.Range("A12").Value
body13 = Body.Range("A13").Value
body14 = Body.Range("A14").Value
body15 = Body.Range("A15").Value
body16 = Body.Range("A16").Value
body18 = Body.Range("A18").Value
body19 = Body.Range("A19").Value
body20 = Body.Range("A20").Value
body22 = Body.Range("A22").Value
body23 = Body.Range("A23").Value
body24 = Body.Range("A24").Value
body26 = Body.Range("A26").Value
body28 = Body.Range("A28").Value
body30 = Body.Range("A30").Value
body32 = Body.Range("A32").Value


Dim OutApp As Object
Dim OutMail As Object

For i = 4 To lr
    'Tao sheet
    Sheets("TBKQ").Select
    Range("B3") = Data.Range("A" & i) 'Gan MNV de chay ham Vlookup
    
    'Tao password de mo file
    Sheets("Data").Select
    pw = Range(colpw & i).Value
    
    'Lay thang dat ten file
    thang = thang_thang & nam_nam
    
    'Tao file xlsx
    Dim wb As Workbook

    Worksheets("TBKQ").Copy
    Set wb = ActiveWorkbook
    ten_file = duongdan & "TBKQ" & i & " - " & thang & ".xlsx"
    wb.SaveAs Filename:=ten_file, Password:=pw

    'Xoa nut chay macro
    ActiveSheet.Buttons.Delete
    Range("K2", "L2").Select
    Selection.ClearContents
    Range("F:G").Select
    Selection.Delete
    'Gan gia tri va dong outline
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False

    ActiveSheet.Outline.ShowLevels RowLevels:=0, ColumnLevels:=1

        
    'Xoa name 'TBKQ'
    For Each NamedRange In ActiveWorkbook.Names
        NamedRange.Delete
    Next NamedRange
    
    'Set PrintArea
    ActiveSheet.PageSetup.PrintArea = "$A$1:$E$61"
    
    'Set password protect sheet
    ActiveSheet.EnableSelection = xlNoSelection
    ActiveSheet.Protect Password:="pw"
    ActiveWorkbook.Protect Password:="pw"
    ActiveWorkbook.Save
    
    wb.Close

    'Gan dia chi gui email va duong dan file dinh kem
    Sheets("Data").Select
    rowemail = i
    dia_chi = Range(colemail & rowemail).Value
    
    dinh_kem = ten_file
    
    'Send email with attachment
    Set OutApp = CreateObject("Outlook.Application")
    OutApp.Session.Logon
    Set OutMail = OutApp.CreateItem(0)

    On Error Resume Next
    With OutMail
        .To = "" & dia_chi & ""
        .Subject = sub1
        .HTMLBody = body1 + "<br />" _
            + "<br />" & body3 _
            + "<br />" & body4 _
            + "<br />" + "<strong>" & body5 + "</strong>" _
            + "<br />" & body6 _
            + "<br />" & body7 _
            + "<br />" & body8 _
            + "<br />" & body9 _
            + "<br />" & body11 _
            + "<br />" & body12 _
            + "<br />" & body13 _
            + "<br />" & body14 _
            + "<br />" & body15 _
            + "<br />" & body16 _
            + "<br />" + "<strong>" & body18 + "</strong>" _
            + "<br />" & body19 _
            + "<br />" & body20 _
            + "<br />" + "<strong>" & body22 + "</strong>" _
            + "<br />" & body23 _
            + "<br />" & body24 _
            + "<br />" & body25 _
            + "<br />" & body26 _
            + "<br />" & body30 + "<br />" _
            + "<br />" + "<strong><span style=""color: rgb(255, 0, 0);"">" & body32 + "</span></strong>" _

        .Attachments.Add ("" & dinh_kem & "")
        '.Display
        .Send
    End With
Next i
    
On Error GoTo 0

Kill duongdan & "*.xlsx"

MsgBox (i - 4 & " Phieu luong da duoc gui.")

Sheets("TBKQ").Select
Range("K2").Value = Range("K2").Value + 1
ActiveSheet.Outline.ShowLevels RowLevels:=0, ColumnLevels:=1

End Sub

