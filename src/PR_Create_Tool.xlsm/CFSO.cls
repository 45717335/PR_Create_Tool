VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFSO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Fso As Object

Public Function folderexists(fdn As String) As Boolean
folderexists = False
If Fso.folderexists(fdn) Then
folderexists = True
End If
End Function

Public Function copyfolder(fdnf As String, fdnt As String) As Boolean
On Error GoTo Errhand
copyfolder = False
If Fso.folderexists(fdnf) And Fso.folderexists(fdnt) Then

If Right(fdnf, 1) = "\" Then fdnf = Left(fdnf, Len(fdnf) - 1)
If Right(fdnt, 1) = "\" Then fdnt = Left(fdnt, Len(fdnt) - 1)


'fdnf = server.MapPath(fdnf)
'  fdnt = server.MapPath(fdnt)
  Call Fso.copyfolder(fdnf, fdnt)
  
copyfolder = True

End If

Exit Function
Errhand:
Msgbox "CFSO.copyfolder ERROR"
End Function

Public Function FileExists(FLFP As String) As Boolean
FileExists = False
If Fso.FileExists(FLFP) Then
FileExists = True
End If
End Function

   
Public Function PickFolder() As String

   
        '** 使用FileDialog对象来选择文件夹
        Dim fd As FileDialog
        Dim strPath As String
       
        Set fd = Application.FileDialog(msoFileDialogFolderPicker)
       
        '** 显示选择文件夹对话框
        If fd.Show = -1 Then        '** 用户选择了文件夹
            strPath = fd.SelectedItems(1)
        Else
            strPath = ""
        End If
        Set fd = Nothing
       PickFolder = strPath
End Function



   
   
Private Sub Class_Initialize()
    '创建类后创建集合
    Set Fso = CreateObject("Scripting.FileSystemObject")
    
End Sub
Private Sub Class_Terminate()
    '类终止后破坏集合
    Set Fso = Nothing
End Sub


  
Public Function CreateFolder(ByVal strPath)
On Error Resume Next
Dim arrPath, pathNum, i, tmpPath, currPath
strPath = Replace(strPath, "\", "/")
arrPath = Split(strPath, "/")
pathNum = UBound(arrPath)
'Set fso = CreateObject("Scripting.FileSystemObject")
If Err <> 0 Then
 Msgbox "ERROR"
End If
For i = 0 To pathNum
   If i = 0 Then
    tmpPath = arrPath(0) & "/"
   Else
    tmpPath = tmpPath & arrPath(i) & "/"
   End If
   currPath = Left(tmpPath, Len(tmpPath) - 1)
   If Not Fso.folderexists(currPath) Then Fso.CreateFolder currPath
Next

End Function

Public Function Datelastmodify(FLFP As String) As Date
 Datelastmodify = Fso.GetFile(FLFP).DateLastModified
End Function

Public Function Userlastmodify(FLFP As String) As String
 Userlastmodify = ""
On Error GoTo Err:

 If FLFP Like "*.xls*" Then
 Dim wb As Workbook
 If open_wb(wb, FLFP) Then
 
 
    
'Dim p As DocumentProperty

 '   For Each p In wb.BuiltinDocumentProperties
'
 '       If Len(p.Value) > 0 Then
'
 '           Debug.Print p.name & "_" & p.Value
'
 '       Else
'
 '           Debug.Print p.name
'
  '      End If
 '
'Next

Userlastmodify = wb.BuiltinDocumentProperties.Item("Last author").Value

 
 
 wb.Saved = True
 wb.Close
 End If
 
 End If
 Exit Function
 
Err:
 Msgbox "FSO ERR:Userlastmodify" & Err.Number
 
End Function


Public Function copy_file(ByVal flfp_from As String, ByVal flfp_to As String) As Boolean

copy_file = False
Dim fdn As String
Dim fln As String
Dim FDN_FROM As String

Dim fln_to As String


fdn = Left(flfp_to, InStrRev(flfp_to, "\"))
FDN_FROM = Left(flfp_from, InStrRev(flfp_from, "\"))
fln = Right(flfp_from, Len(flfp_from) - Len(FDN_FROM))

If Right(flfp_to, 1) = "\" Then
fln_to = fln
Else
fln_to = ""
End If



CreateFolder fdn


'Dim fso As Object
'Set fso = CreateObject("Scripting.FileSystemObject")
'Dim fd As Object
'Set fd = FSO.GetFolder(FDN_BOM)
          
If Fso.FileExists(flfp_from) = True Then
        Fso.CopyFile flfp_from, flfp_to & fln_to, True
        copy_file = True
End If



            
End Function

