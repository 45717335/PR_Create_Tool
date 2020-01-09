Attribute VB_Name = "FSO_SENIOR"
Option Explicit

'使用了 CFSO,OneKeyCls

Function Record_file_in_folder(mokc As OneKeyCls, Optional FDN_root As String = "", Optional FLN_include As String = "") As String
On Error Resume Next

Dim b_continue As Boolean
b_continue = True
Dim rec_b As Boolean

Dim i As Long

'mokc.item("PARAP").item("FDN_root")
'mokc.item("PARAP").item("FDN_current")
'mokc.item("PARA").item("FLN_include")

'合法性检测

If mokc.Item("PARA") Is Nothing And FDN_root <> "" And FLN_include <> "" Then
mokc.Add "PARA", "PARA"
mokc.Item("PARA").Add FDN_root, "FDN_root"
mokc.Item("PARA").Add FDN_root, "FDN_current"
mokc.Item("PARA").Add "FLN_include", "FLN_include"
mokc.Item("PARA").Item("FLN_include").Add FLN_include
End If


If mokc.Item("PARA").Item("FDN_root") Is Nothing Then b_continue = False
If mokc.Item("PARA").Item("FDN_current") Is Nothing Then b_continue = False
If mokc.Item("PARA").Item("FLN_include") Is Nothing Then


mokc.Item("PARA").Add "FLN_include", "FLN_include"
mokc.Item("PARA").Item("FLN_include").Add ".tif"
mokc.Item("PARA").Item("FLN_include").Add ".xls"

End If

If b_continue = False Then
Record_file_in_folder = "ERROR:Record_file_in_folder"
Msgbox Record_file_in_folder
Exit Function
End If
'合法性检测


If mokc.Item("FILE") Is Nothing Then mokc.Add "FILE", "FILE"
FDN_root = mokc.Item("PARA").Item("FDN_root").Key



Dim Fso As Object
Set Fso = CreateObject("Scripting.FileSystemObject")
Dim fd As Object
Set fd = Fso.GetFolder(mokc.Item("PARA").Item("FDN_current").Key)
Dim fl As Object
Dim sfd As Object
Dim cur_i As Integer
cur_i = 1
Dim temp_s As String
Dim i_curr As Long



For Each fl In fd.Files


rec_b = True
temp_s = fl.Path
If Len(temp_s) > Len(FDN_root) Then
temp_s = Right(temp_s, Len(temp_s) - Len(FDN_root))
End If

If rec_b Then

For i = 1 To mokc.Item("PARA").Item("FLN_include").Count
rec_b = False

If InStr(temp_s, mokc.Item("PARA").Item("FLN_include").Item(i).Key) > 0 Then
rec_b = True
Exit For
End If
Next
End If




If rec_b Then
i_curr = mokc.Item("FILE").Count
i_curr = i_curr + 1
mokc.Item("FILE").Add CStr(i_curr), CStr(i_curr)

mokc.Item("FILE").Item(CStr(i_curr)).Add fl.Name, "FLN"
mokc.Item("FILE").Item(CStr(i_curr)).Add CStr(fl.Size), "SIZE"
mokc.Item("FILE").Item(CStr(i_curr)).Add mokc.Item("PARA").Item("FDN_current").Key, "FDN"
mokc.Item("FILE").Item(CStr(i_curr)).Add Format(fl.DateLastModified, "YYYY-MM-DD HH:MM:SS"), "DATE"

End If



Next fl


'For Each sfd In fd.SubFolders
'mokc.Item("PARA").Item("FDN_current").Key = sfd.Path
'Record_file_in_folder mokc
'Next sfd

End Function


Function mokc_read_ws(mokc As OneKeyCls, ws As Worksheet, Optional key_i1 As Integer = 1, Optional key_i2 As Integer = 0)
'将电子表格中的内容读入mokc，电子表格的 名称和，A1，B1...单元格的内容是关键字
If key_i2 = 0 Then key_i2 = key_i1


Dim wsn As String
wsn = ws.Name
If Not (mokc.Item(wsn) Is Nothing) Then
mokc.Remove wsn
End If
mokc.Add wsn, wsn
Dim i As Integer
Dim i_last As Integer
Dim j As Integer

i_last = ws.UsedRange.Columns.Count
Dim temp_s1 As String
Dim temp_s2 As String
Dim temp_s3 As String
Dim temp_s4 As String
Dim temp_s5 As String

mokc.Item(wsn).Add "HEAD", "HEAD"
mokc.Item(wsn).Add "BODY", "BODY"
mokc.Item(wsn).Add "KEY", "KEY"
mokc.Item(wsn).Item("KEY").Add CStr(key_i1), "KEY1"
mokc.Item(wsn).Item("KEY").Add CStr(key_i2), "KEY2"



For i = 1 To i_last
temp_s1 = Trim(ws.Cells(1, i))
If Len(temp_s1) > 0 Then
If mokc.Item(wsn).Item("HEAD").Item(temp_s1) Is Nothing Then
mokc.Item(wsn).Item("HEAD").Add temp_s1, temp_s1
mokc.Item(wsn).Item("HEAD").Item(temp_s1).Add CStr(i), CStr(i)
End If
End If
Next

i_last = ws.UsedRange.Rows.Count

For i = 2 To i_last

    temp_s1 = Trim(ws.Cells(i, key_i1))
    temp_s2 = Trim(ws.Cells(i, key_i2))
    If Len(temp_s2) = 0 Then temp_s2 = temp_s1
    
    If Len(temp_s1) > 0 And Len(temp_s2) > 0 Then
    If mokc.Item(wsn).Item("BODY").Item(temp_s1) Is Nothing Then mokc.Item(wsn).Item("BODY").Add temp_s1, temp_s1
    If mokc.Item(wsn).Item("BODY").Item(temp_s1).Item(temp_s2) Is Nothing Then mokc.Item(wsn).Item("BODY").Item(temp_s1).Add temp_s2, temp_s2
    temp_s3 = CStr(i)
    If mokc.Item(wsn).Item("BODY").Item(temp_s1).Item(temp_s2).Item("#ROW") Is Nothing Then mokc.Item(wsn).Item("BODY").Item(temp_s1).Item(temp_s2).Add temp_s3, temp_s3
    For j = 1 To mokc.Item(wsn).Item("HEAD").Count
    temp_s4 = mokc.Item(wsn).Item("HEAD").Item(j).Key
    temp_s5 = ws.Cells(i, CInt(mokc.Item(wsn).Item("HEAD").Item(j).Item(1).Key))
    mokc.Item(wsn).Item("BODY").Item(temp_s1).Item(temp_s2).Item(temp_s3).Add temp_s5, temp_s4
    Next
    




End If





Next






End Function


Function mokc_read_ws_A(mokc As OneKeyCls, ws As Worksheet, Optional i As Long = 0) As Boolean
'读取工作表中的一行到 mokc中
Dim temp_s1 As String, temp_s2 As String, temp_s3 As String, temp_s4 As String, temp_s5 As String

Dim key_i1 As Integer, key_i2 As Integer

Dim wsn As String
wsn = ws.Name
Dim j As Integer


If i = 0 Then i = ws.UsedRange.Rows(ws.UsedRange.Rows.Count).row
If mokc.Item(ws.Name).Item("BODY") Is Nothing Or mokc.Item(ws.Name).Item("HEAD") Is Nothing Then
mokc_read_ws_A = False
Msgbox "Error! mokc_read_ws_A " & Chr(10) & ws.Name
End If

    
    temp_s1 = Trim(ws.Cells(i, CInt(mokc.Item(ws.Name).Item("KEY").Item("KEY1").Key)))
    temp_s2 = Trim(ws.Cells(i, CInt(mokc.Item(ws.Name).Item("KEY").Item("KEY2").Key)))
    If Len(temp_s2) = 0 Then temp_s2 = temp_s1
    
    If Len(temp_s1) > 0 And Len(temp_s2) > 0 Then
    If mokc.Item(wsn).Item("BODY").Item(temp_s1) Is Nothing Then mokc.Item(wsn).Item("BODY").Add temp_s1, temp_s1
    If mokc.Item(wsn).Item("BODY").Item(temp_s1).Item(temp_s2) Is Nothing Then mokc.Item(wsn).Item("BODY").Item(temp_s1).Add temp_s2, temp_s2
    temp_s3 = CStr(i)
    If mokc.Item(wsn).Item("BODY").Item(temp_s1).Item(temp_s2).Item(temp_s3) Is Nothing Then mokc.Item(wsn).Item("BODY").Item(temp_s1).Item(temp_s2).Add temp_s3, temp_s3
    For j = 1 To mokc.Item(wsn).Item("HEAD").Count
    temp_s4 = mokc.Item(wsn).Item("HEAD").Item(j).Key
    temp_s5 = ws.Cells(i, CInt(mokc.Item(wsn).Item("HEAD").Item(j).Item(1).Key))
    mokc.Item(wsn).Item("BODY").Item(temp_s1).Item(temp_s2).Item(temp_s3).Add temp_s5, temp_s4
    Next
    End If
    


    
End Function
