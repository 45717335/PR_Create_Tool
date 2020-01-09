Attribute VB_Name = "Excel_VBA"
Option Explicit
'
'Function open_wb(ByRef wb As Workbook, ByVal flfp As String) As Boolean
'Function ws_exist(wb As Workbook, wsn As String) As Boolean
'Function get_ws(ByRef wb As Workbook, ByVal wsname As String) As Worksheet
'Function add_comm(ByVal comm_s As String, ws1 As Worksheet, ByVal h_i As Integer, ByVal l_i As Integer, ByVal visiable As Boolean) As Boolean
'Function open_wb2(ByRef wb As Workbook, ByVal flfp As String) As Boolean
'Function Str_TO_Num(in_s As String, ByRef out_i As Integer) As Boolean
'Function append_ws(ByRef ws As Worksheet, ByVal a As String, ByVal A_val) As Integer
'Function GetColName(ByVal intCol As Long) As String
'Function check_cell(rg As Range, s_desir As String) As Boolean                (2017-06-14)


Function check_cell(rg As Range, s_desir As String) As Boolean
'检查单元格是否为指定内容，如果不是，标红，并在在备注提示
Dim s_temp As String
s_temp = rg.Text
If s_temp = s_desir Then
check_cell = True
Else
check_cell = False
If rg.Comment Is Nothing Then
rg.AddComment
End If
rg.Comment.Text Text:=s_desir
rg.Comment.Visible = True
End If
End Function

Function open_wb(ByRef wb As Workbook, ByVal FLFP As String) As Boolean
'==========================================================
'Open File(*.xls*):  Microsoft Excel
'==========================================================
open_wb = False

Dim i As Integer
Dim fln, flp As String
fln = Right(FLFP, Len(FLFP) - InStrRev(FLFP, "\"))
flp = Left(FLFP, Len(FLFP) - Len(fln))
Dim temp_b As Boolean
temp_b = False
For i = 1 To Workbooks.Count
If Workbooks(i).Name = fln Then
temp_b = True
Set wb = Workbooks(i)
Exit For
End If
Next
If temp_b = False Then
If Dir(flp & fln) <> "" Then

On Error GoTo Error1:
Set wb = Workbooks.Open(flp & fln)

temp_b = True
End If
End If
open_wb = temp_b
Exit Function
Error1:
    Msgbox "open_wb function:" + Err.Description
    Err.Clear
    Exit Function
    
End Function

Function ws_exist(wb As Workbook, wsn As String) As Boolean
'本函数用于判断工作簿中是否存在指定工作表
Dim ws As Worksheet
ws_exist = False
For Each ws In wb.Worksheets
If ws.Name = wsn Then
ws_exist = True
Exit Function
End If
Next
End Function


Function get_ws(ByRef wb As Workbook, ByVal wsname As String) As Worksheet
On Error GoTo ERRORHAND
Dim i As Integer
Dim havewsT As Boolean
havewsT = False
For i = 1 To wb.Worksheets.Count
If wb.Worksheets(i).Name = wsname Then
Set get_ws = wb.Worksheets(i)
havewsT = True
End If
Next
If havewsT = False Then
wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count)).Name = wsname
Set get_ws = wb.Worksheets(wsname)
End If
Exit Function
ERRORHAND:
If Err.Number <> 0 Then Msgbox "get_ws function: " + Err.Description
Err.Clear
End Function

Function add_comm(ByVal comm_s As String, ws1 As Worksheet, ByVal h_i As Integer, ByVal l_i As Integer, ByVal visiable As Boolean) As Boolean
On Error GoTo ERRORHAND
If ws1.Cells(h_i, l_i).Comment Is Nothing Then
    ws1.Cells(h_i, l_i).AddComment
End If
ws1.Cells(h_i, l_i).Comment.Text Text:=comm_s
ws1.Cells(h_i, l_i).Comment.Visible = visiable
Exit Function
ERRORHAND:
If Err.Number <> 0 Then Msgbox "get_ws function: " + Err.Description
Err.Clear
End Function


Function add_comment(ByVal comm_s As String, tar_rg As Range) As Boolean
On Error GoTo ERRORHAND

If tar_rg.Comment Is Nothing Then
    tar_rg.AddComment
End If
tar_rg.Comment.Text Text:=comm_s
tar_rg.Comment.Visible = True
Exit Function
ERRORHAND:
If Err.Number <> 0 Then Msgbox "get_ws function: " + Err.Description
Err.Clear
End Function

Function open_wb2(ByRef wb As Workbook, ByVal FLFP As String) As Boolean
'==========================================================
'在新窗口中打开 workbook
'==========================================================
open_wb2 = False

   Dim app As Object
   Set app = CreateObject("Excel.application")
   app.Visible = True
   
   
Dim i As Integer
Dim fln, flp As String
fln = Right(FLFP, Len(FLFP) - InStrRev(FLFP, "\"))
flp = Left(FLFP, Len(FLFP) - Len(fln))
Dim temp_b As Boolean
temp_b = False
For i = 1 To app.Workbooks.Count
If app.Workbooks(i).Name = fln Then
temp_b = True
Set wb = app.Workbooks(i)
Exit For
End If
Next
If temp_b = False Then
If Dir(flp & fln) <> "" Then

On Error GoTo Error1:
Set wb = app.Workbooks.Open(flp & fln)

temp_b = True
End If
End If
open_wb2 = temp_b
Exit Function
Error1:
    Msgbox "open_wb2 function:" + Err.Description
    Err.Clear
    Exit Function
    
End Function


Function Close_wb2(ByRef wb As Workbook) As Boolean
'==========================================================
'在新窗口中打开 workbook
'==========================================================
On Error GoTo ERRORHAND
Dim app As Object
Set app = wb.Application
If wb.Application.Workbooks.Count = 1 Then
wb.Close
app.Quit
Set app = Nothing
End If
Exit Function
ERRORHAND:
Msgbox "Close_wb2 function:" + Err.Description
Err.Clear
End Function


Function Str_TO_Num(in_s As String, ByRef out_i As Integer) As Boolean
'本函数用于字符串转数字
On Error GoTo ERRORHAND
Str_TO_Num = True
out_i = CInt(in_s)
Exit Function
ERRORHAND:
Str_TO_Num = False
'MsgBox "Str_TO_Num:" + Err.Description
Err.Clear
End Function

Function Str_TO_Dbl(in_s As String, ByRef out_dbl As Double) As Boolean
'本函数用于字符串转数字
On Error GoTo ERRORHAND
Str_TO_Dbl = True
out_dbl = CDbl(in_s)
Exit Function
ERRORHAND:
Str_TO_Dbl = False
'MsgBox "Str_TO_Num:" + Err.Description
Err.Clear
End Function





Function append_ws(ByRef ws As Worksheet, ByVal a As String, ByVal A_val) As Integer
append_ws = 0
Dim lastrow As Integer
lastrow = ws.UsedRange.Rows(ws.UsedRange.Rows.Count).row
ws.Range(a & lastrow + 1) = A_val
append_ws = lastrow + 1

End Function

Function GetColName(ByVal intCol As Long) As String
'列号转列名
If InStr(CStr(Application.Version), "11") > 0 And intCol >= 1 And intCol <= 256 Then
    GetColName = Split(Workbooks(1).Worksheets(1).Cells(1, intCol).Address, "$")(1)
ElseIf InStr(CStr(Application.Version), "12") > 0 And intCol >= 1 And intCol <= 16384 Then
    GetColName = Split(Workbooks(1).Worksheets(1).Cells(1, intCol).Address, "$")(1)

ElseIf InStr(CStr(Application.Version), "14") > 0 And intCol >= 1 And intCol <= 16384 Then
    GetColName = Split(Workbooks(1).Worksheets(1).Cells(1, intCol).Address, "$")(1)

Else

    GetColName = "Error"
End If
End Function

Function my_cint(str1 As String) As Integer
On Error GoTo ErrorH
my_cint = CInt(str1)
Exit Function
ErrorH:
my_cint = 0
End Function

Function my_msgbox(str1 As String, Optional tf As Boolean = True) As VbMsgBoxResult
my_msgbox = Msgbox(str1, tf)
End Function


Function rg_scrowll(rg As Range)
    '滚动到指定单元格
    Dim ws As Worksheet
    Dim i_r As Integer, i_c As Integer, i_wr As Integer, I_wc As Integer
    Set ws = rg.Parent
    ws.Activate
    i_r = rg.row
    i_c = rg.Column
    i_wr = ActiveWindow.VisibleRange.row
    I_wc = ActiveWindow.VisibleRange.Column
    If i_r - i_wr = 0 Then
    Else
        ActiveWindow.SmallScroll Down:=i_r - i_wr
    End If
    If i_c - I_wc = 0 Then
    Else
        ActiveWindow.SmallScroll ToRight:=i_c - I_wc
    End If
End Function

