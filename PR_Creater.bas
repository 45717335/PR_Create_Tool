Attribute VB_Name = "PR_Creater"
Option Explicit

Sub PR_Creater()
Attribute PR_Creater.VB_ProcData.VB_Invoke_Func = "r\n14"
back_followinglist
'MsgBox "WinShuttle Can Use again! Please UPload to SAP after Create PR!"
 If Enagble_addins("PPPM") Then
 End If
 
Dim wb_pr As Workbook

Dim mfso As New CFSO
Dim mokc As New OneKeyCls
mokc.Add "FL", "FL"
mokc.Add "PR", "PR"
mokc.Item("PR").Add "FLFP_TEMPLATE", "FLFP_TEMPLATE"
mokc.Item("PR").Add "OEM_NAME", "OEM_NAME"
mokc.Add "WS_PartSingle", "WS_PartSingle"
mokc.Item("WS_PartSingle").Add "WS_HEAD", "WS_HEAD"
mokc.Item("WS_PartSingle").Add "WS_BODY", "WS_BODY"


mokc.Item("FL").Add "FLN", "FLN"
mokc.Item("FL").Add "FDN", "FDN"

mokc.Item("PR").Add "PRN_LAST", "PRN_LAST"

mokc.Item("FL").Item("FDN").Add "FDNPR", "FDNPR"

mokc.Item("FL").Add "CUR_PR_NUM", "CUR_PR_NUM"



'20181129 ȥ����Ӧ�������е������ַ�
Dim mokc_manu As New OneKeyCls
If ws_exist(Workbooks("PR_Create_Tool.xlsm"), "MANUFATURE") Then
mokc_read_ws mokc_manu, Workbooks("PR_Create_Tool.xlsm").Worksheets("MANUFATURE"), 1, 1
End If
'20181129 ȥ����Ӧ�������е������ַ�


'mokc_read_ws(mok

'=======================
'�ҵ��Ѿ��򿪵Ļ�е���ٱ�
Dim b_c As Boolean
Dim str1 As String
Dim str2 As String
Dim str3 As String
Dim str4 As String

Dim s_t1 As String
Dim s_t2 As String
Dim s_t3 As String
Dim s_t4 As String


Dim wb_fl As Workbook
Dim i As Long
Dim i_last As Long
Dim i_PRN_LAST As Integer
Dim s_PRN_LAST As String
Dim fln As String
Dim fdn As String
Dim ws_partsingle As Worksheet

Dim j As Long
Dim j_last As Long

'��ȡPartSingle����ĵ�Ԫ��
Dim POS As String
Dim QTY As String
Dim UNIT As String
Dim ItemName As String
Dim OEM_ID As String
Dim OEM_NAME As String
Dim TKID_SUBASS As String
Dim TKID_STATION As String
Dim PA_Index As String
Dim R_DATE As String
Dim E_DATE As String
Dim MATERIAL As String
Dim STANDARD As String
Dim DIMENSION As String
Dim dbl_qty As Double



'��ȡPartSingle����ĵ�Ԫ��




b_c = False
For i = 1 To Workbooks.Count
str1 = Workbooks(i).Name
If str1 Like "CN.*Mechanics*Following*" Or str1 Like "CN.*Following*Mechanics*" Then
Set wb_fl = Workbooks(i)
b_c = True
Exit For
End If
Next
If Not (wb_fl Is Nothing) Then
Msgbox "Following list to create PR��" & Chr(10) & wb_fl.Name
If wb_fl.ReadOnly = True Then
Msgbox "ReadOnly Following list Can Not create PR!"
wb_fl.Close
Exit Sub
End If
Else
Msgbox "Please open following list in the first!�� CN.*Mechanics*Following.xlsm "
Exit Sub
End If
'�ҵ��Ѿ��򿪵Ļ�е���ٱ�
'=======================



'20190226 ������ȡBOM�� PartSingle
If Read_Main_to_PS(wb_fl) = True Then
'��� �ж�ȡ��������ֹ ����PR���̣��ù���ʦ��� ��Main�����PartSingle������
Exit Sub
End If
wb_fl.Save
'20190226 ������ȡBOM�� PartSingle





mokc.Item("FL").Item("FLN").Key = wb_fl.Name
mokc.Item("FL").Item("FDN").Key = wb_fl.Path
If Right(mokc.Item("FL").Item("FDN").Key, 1) <> "\" Then
mokc.Item("FL").Item("FDN").Key = mokc.Item("FL").Item("FDN").Key & "\"
End If
'=======================
'Checking  �ļ����Ƿ���� (mokc.Item("FL").Item("FDN").Item("FDNPR").Key)
fdn = mokc.Item("FL").Item("FDN").Key & "PAX_Mechanical&H&P"
mokc.Item("FL").Item("FDN").Item("FDNPR").Key = fdn
If mfso.folderexists(fdn) = False Then
If Msgbox("Folder does not exist!:" & Chr(10) & fdn & Chr(10) & "Create press OK.", vbOKCancel) = vbOK Then
mfso.CreateFolder fdn
Msgbox "Folder to store PR was created! :" & Chr(10) & mokc.Item("FL").Item("FDN").Key & "PAX_Mechanical&H&P" & Chr(10) & "Please put a PR in that folder ��Defalte  PR Number 0002��"
Else
Exit Sub
End If
End If
'Checking  �ļ����Ƿ���� (mokc.Item("FL").Item("FDN").Item("FDNPR").Key)
'=======================







'=======================
'Get_Last PAX Number

i_PRN_LAST = 1
s_PRN_LAST = "0001"
Record_file_in_folder mokc.Item("FL").Item("FDN").Item("FDNPR"), mokc.Item("FL").Item("FDN").Item("FDNPR").Key, ".xlsm"
For i = 1 To mokc.Item("FL").Item("FDN").Item("FDNPR").Item("FILE").Count
fln = mokc.Item("FL").Item("FDN").Item("FDNPR").Item("FILE").Item(i).Item("FLN").Key
fdn = mokc.Item("FL").Item("FDN").Item("FDNPR").Item("FILE").Item(i).Item("FDN").Key
If fdn = mokc.Item("FL").Item("FDN").Item("FDNPR").Key Then
    If fln Like "P?####*.xlsm" Then
    str1 = Mid(fln, 3, 4)
        If CInt(str1) > CInt(s_PRN_LAST) Then
        s_PRN_LAST = str1
        i_PRN_LAST = CInt(s_PRN_LAST)
       
        
        
        End If
    ElseIf fln Like "PAX###*.xlsm" Then
    
    
        str1 = Mid(fln, 4, 3)
        If CInt(str1) > CInt(s_PRN_LAST) Then
        s_PRN_LAST = str1
        i_PRN_LAST = CInt(s_PRN_LAST)
       
        
        End If
    Else
    
    If Not (fln Like "MO*") And Not (fln Like "DM*") And Not (fln Like "~$*") Then
    Msgbox fln
    End If
    
    
    End If
End If
Next
i_PRN_LAST = i_PRN_LAST + 1
s_PRN_LAST = CStr(i_PRN_LAST)
If Len(s_PRN_LAST) < 4 Then
s_PRN_LAST = Left("000", 4 - Len(s_PRN_LAST)) & s_PRN_LAST
End If
mokc.Item("PR").Item("PRN_LAST").Key = s_PRN_LAST
'Get_Last PAX Number
'=======================




'=====================
'��ȡ Part_single ���
b_c = True ' b_c = True  Part_Single ��ȡ�ɹ���b_c = False  Part_Single ��ȡʧ��
mokc.Item("WS_PartSingle").Add "M_C_P", "M_C_P"
'Record the pype: Controls or Mechanics or Pneumatics
'������ Parts_Single,ʧ��
If b_c Then
If ws_exist(wb_fl, "Parts_Single") = False Then
b_c = False
Msgbox "Following list ,Does not exist{ Parts_Single}"
Else
Set ws_partsingle = wb_fl.Worksheets("Parts_Single")
End If
End If
'�޷��ж��� �� or �� or �� ʧ��
If b_c Then
If InStr(wb_fl.Name, "Mechanics") > 0 Then
mokc.Item("WS_PartSingle").Item("M_C_P").Key = "Mechanics"
ElseIf InStr(wb_fl.Name, "Controls") > 0 Then
mokc.Item("WS_PartSingle").Item("M_C_P").Key = "Controls"
ElseIf InStr(wb_fl.Name, "Pneumatics") > 0 Then
mokc.Item("WS_PartSingle").Item("M_C_P").Key = "Pneumatics"
Else
b_c = False
Msgbox wb_fl.Name & Chr(10) & "File name must contain one of ��Mechanics or  Controls or Pneumatics"
End If
End If
'�ж� Part_Single ��ʽ�Ƿ���Ԥ�ø�ʽ
If b_c Then
If mokc.Item("WS_PartSingle").Item("M_C_P").Key = "Mechanics" Then

'Template 1
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 1, "Pos.", "POS") Then b_c = False
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 3, "Qty", "QTY") Then b_c = False
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 4, "Base Unit", "UNIT") Then b_c = False
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 6, "Matl. Descrip.", "ItemName") Then b_c = False
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 7, "Material No.", "TKID") Then b_c = False
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 8, "SPI", "SPI") Then b_c = False
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 11, "Manuf.Part.No.", "OEM_ID") Then b_c = False
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 12, "Basic Material", "MATERIAL") Then b_c = False
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 13, "Matl. Standard", "STANDARD") Then b_c = False
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 14, "Size/Dimension", "DIMENSION") Then b_c = False
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 15, "Manufacturer", "OEM_NAME") Then b_c = False
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 16, "Sub-Assy._Number", "TKID_SUBASS") Then b_c = False
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 17, "Station_Number", "TKID_STATION") Then b_c = False
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 18, "PA_Index", "PA_Index") Then If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 18, "PA_Number", "PA_Index") Then b_c = False
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 19, "Release_date", "R_DATE") Then b_c = False
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 20, "Expect Week", "E_DATE") Then b_c = False
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 21, "description", "DESC") Then If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 21, "Description", "DESC") Then b_c = False
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 23, "MO ID", "MO ID") Then b_c = False
    If b_c = False Then Msgbox "Part_Single Table head unknown!"
    
    
End If
End If



If b_c Then

i_last = ws_partsingle.UsedRange.Rows(ws_partsingle.UsedRange.Rows.Count).row
For i = 8 To i_last
'��ȡ������Ϊ�㣬PA_Index Ϊ�յ���




PA_Index = Trim(ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("PA_Index").Item(1).Key)))
QTY = Trim(ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("QTY").Item(1).Key)))


If Len(PA_Index) = 0 And Len(QTY) > 0 Then
OEM_NAME = Trim(ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("OEM_NAME").Item(1).Key)))
TKID_STATION = Trim(ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("TKID_STATION").Item(1).Key)))
dbl_qty = 0
Str_TO_Dbl QTY, dbl_qty
If Not (dbl_qty > 0) Then
Msgbox "Qty Must >0 ��row number ��" & i
b_c = False
End If
If b_c Then
If Len(OEM_NAME) = 0 Then
Msgbox "Name of muaufature can not be empty  ro number ��" & i
b_c = False
End If





If b_c Then


If OEM_NAME = "TKSY" And InStr(ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("TKID").Item(1).Key)), "SP.00") > 0 Then
Msgbox "?.?????.???.SP.00  No Pr , Row: " & i
ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("PA_Index").Item(1).Key)) = "SP.NA"

Else

mokc.Item("WS_PartSingle").Item("WS_BODY").Add CStr(i), CStr(i)
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add Trim(ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("POS").Item(1).Key))), "POS"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add Trim(ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("QTY").Item(1).Key))), "QTY"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add Trim(ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("UNIT").Item(1).Key))), "UNIT"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add Trim(ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("ItemName").Item(1).Key))), "ItemName"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add Trim(ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("TKID").Item(1).Key))), "TKID"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add Trim(ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("SPI").Item(1).Key))), "SPI"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add Trim(ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("OEM_ID").Item(1).Key))), "OEM_ID"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add Trim(ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("MATERIAL").Item(1).Key))), "MATERIAL"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add Trim(ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("STANDARD").Item(1).Key))), "STANDARD"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add Trim(ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("DIMENSION").Item(1).Key))), "DIMENSION"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add Trim(ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("OEM_NAME").Item(1).Key))), "OEM_NAME"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add Trim(ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("TKID_SUBASS").Item(1).Key))), "TKID_SUBASS"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add Trim(ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("TKID_STATION").Item(1).Key))), "TKID_STATION"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add Trim(ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("PA_Index").Item(1).Key))), "PA_Index"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add Trim(ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("R_DATE").Item(1).Key))), "R_DATE"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add Trim(ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("MO ID").Item(1).Key))), "MO ID"
E_DATE = format_date_DDMMYYYY(ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("E_DATE").Item(1).Key)))
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add E_DATE, "E_DATE"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("DESC").Item(1).Key)), "DESC"

End If

End If
End If
End If
Next
End If
'��ȡ Part_single ���
'=====================

If b_c = False Then
Msgbox "Can not read  Part_Single fail to create PR!"
Exit Sub
End If





'=====================
'���PR ģ���Ƿ���ڣ������ڿ���һ�ݵ���ǰĿ¼
'Z:\24_Temp\PA_Logs\V1.1\TEMPLATE\010c1612_Purchase Requisition.xlsm
If mfso.FileExists(mokc.Item("FL").Item("FDN").Key & "\010c1612_Purchase Requisition.xlsm") Then
Else
 If mfso.FileExists("Z:\24_Temp\PA_Logs\V1.1\TEMPLATE\010c1612_Purchase Requisition.xlsm") Then
 mfso.copy_file "Z:\24_Temp\PA_Logs\V1.1\TEMPLATE\010c1612_Purchase Requisition.xlsm", mokc.Item("FL").Item("FDN").Key & "\010c1612_Purchase Requisition.xlsm"
 Else
 Msgbox "�� PRģ�壺Z:\24_Temp\PA_Logs\V1.1\TEMPLATE\010c1612_Purchase Requisition.xlsm"
 b_c = False
 End If
End If
mokc.Item("PR").Item("FLFP_TEMPLATE").Key = mokc.Item("FL").Item("FDN").Key & "\010c1612_Purchase Requisition.xlsm"
'���PR ģ���Ƿ���ڣ������ڿ���һ�ݵ���ǰĿ¼
'=====================
If b_c = False Then
Msgbox "PR template does not exist ! can not create  PR"
Exit Sub
End If



'==============================
'���ģ������Ŀ����Ŀ���Ƿ���ڣ�������Ҫ������
If open_wb(wb_pr, mokc.Item("PR").Item("FLFP_TEMPLATE").Key) Then
str1 = wb_pr.Worksheets(1).Range("G7")
Do While Len(str1) = 0
str1 = InputBox("Please input project number", "PR template", "CN.#######")
If str1 = "CN.#######" Then str1 = ""
Loop
wb_pr.Worksheets(1).Range("G7") = str1
str1 = wb_pr.Worksheets(1).Range("M7")
Do While Len(str1) = 0
str1 = InputBox("Please input project name", "PR template")
Loop
wb_pr.Worksheets(1).Range("M7") = str1
wb_pr.Save
wb_pr.Saved = True
wb_pr.Close
Else
b_c = False
End If
'���ģ������Ŀ����Ŀ���Ƿ���ڣ�������Ҫ������
'==============================
If b_c = False Then
Msgbox "PR ģ�� ����Ŀ���ƣ�����Ŀ��δ��д"
Exit Sub
End If






'============================
'MO ��飬�����MO ����MO
Dim i_moid As Integer
i_moid = 0
For i = 1 To mokc.Item("WS_PartSingle").Item("WS_BODY").Count
str1 = UNIT_check(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Item("MO ID").Key)
If Len(str1) > 0 Then
If my_cint(str1) = 0 And i_moid > 0 Then
Msgbox "MO ���Ӻ� PX ���ӱ���ֿ��£�"
b_c = False
Exit For
End If
If i_moid < my_cint(str1) Then
i_moid = my_cint(str1)
If i_moid > 999 Then
Msgbox "MO ��ű���С��999!"
b_c = False
Exit For
End If
End If
End If
Next

'MO ��飬�����MO ����MO
'============================
If b_c = False Then
Msgbox "MO Check Error�� " & Chr(10) & "Row:" & mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Key
Exit Sub
End If


'===================
'�����MO�����޸�
If i_moid > 0 Then
mokc.Item("PR").Item("PRN_LAST").Key = CStr(i_moid * 10)
mokc.Item("PR").Item("PRN_LAST").Key = Right("000" & mokc.Item("PR").Item("PRN_LAST").Key, 4)
End If
'�����MO�����޸�
'===================




'============================
'��λ ���
For i = 1 To mokc.Item("WS_PartSingle").Item("WS_BODY").Count
str1 = UNIT_check(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Item("UNIT").Key)
If Len(str1) = 0 Then
b_c = False
Exit For
Else
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Item("UNIT").Key = str1
End If
Next

'��λ ���
'============================
If b_c = False Then
Msgbox "Unit Check fail." & Chr(10) & "Row:" & mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Key
Exit Sub
End If





'============================2019 12 09
'��λ�ż��  ֻ���� CN.??????.???, �� "D.?????.???"
For i = 1 To mokc.Item("WS_PartSingle").Item("WS_BODY").Count
str1 = mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Item("TKID_STATION").Key

If Not (str1 Like "CN.??????*" Or str1 Like "D.?????.???*") Then
b_c = False
Exit For
End If
Next
'��λ�ż��  ֻ���� CN.??????.???, �� "D.?????.???"
'============================2019 12 09
If b_c = False Then
Msgbox "Station Number Check fail Must Be : CN.??????.??? OR  D.?????.???��" & Chr(10) & "Row:" & mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Key
Exit Sub
End If






'============================
'�������� ���
For i = 1 To mokc.Item("WS_PartSingle").Item("WS_BODY").Count
str1 = mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Item("E_DATE").Key
If Len(str1) = 0 Then
b_c = False
Exit For
Else
End If
Next

'�������� ���
'============================
If b_c = False Then
Msgbox "Date receive error:" & Chr(10) & "Row:" & mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Key
Exit Sub
End If





'============================
'��Ӧ������ ��鼰����
For i = 1 To mokc.Item("WS_PartSingle").Item("WS_BODY").Count
str1 = OEM_NAME_check(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Item("OEM_NAME").Key)
If Len(str1) = 0 Then
b_c = False
Exit For
End If
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Item("OEM_NAME").Key = str1


If mokc.Item("PR").Item("OEM_NAME").Item(str1) Is Nothing Then
mokc.Item("PR").Item("OEM_NAME").Add str1, str1
mokc.Item("PR").Item("OEM_NAME").Item(str1).Add CStr(i), CStr(i)


Else
mokc.Item("PR").Item("OEM_NAME").Item(str1).Add CStr(i), CStr(i)
End If
Next
'��Ӧ������ ��鼰����
'============================




If b_c = False Then
Msgbox "��ȡ ��Ӧ�̼�� ʧ�ܣ��޷�����PR" & Chr(10) & "�к�:" & mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Key
Exit Sub
End If


'============================
'��׼�� �� ���� PA ����ֿ���
If mokc.Item("PR").Item("OEM_NAME").Count > 1 Then
If Not (mokc.Item("PR").Item("OEM_NAME").Item("NA") Is Nothing) Then

Msgbox "�кţ�" & mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item("NA").Item(1).Key)).Key
b_c = False

End If
ElseIf Not (mokc.Item("PR").Item("OEM_NAME").Item("NA") Is Nothing) Then
Msgbox "NA DoNot Use This!"
b_c = False
Exit Sub
End If
'��׼�� �� ���� PA ����ֿ���
'============================
If b_c = False Then
Msgbox "Screws (NA)��Must create PR in separate."
Exit Sub
End If



'=====================
'N/A ֱ�� ��д N/A
'��Ӧ�����Ʋ��ܰ��������ַ�
If mokc.Item("PR").Item("OEM_NAME").Count > 1 Then
If Not (mokc.Item("PR").Item("OEM_NAME").Item("N/A") Is Nothing) Then
For i = 1 To mokc.Item("PR").Item("OEM_NAME").Item("N/A").Count
ws_partsingle.Cells(CInt(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item("N/A").Item(i).Key)).Key), CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("PA_Index").Item(1).Key)) = "N/A"
Next
End If
End If

If Not (mokc.Item("PR").Item("OEM_NAME").Item("N/A") Is Nothing) Then mokc.Item("PR").Item("OEM_NAME").Remove "N/A"

For i = 1 To mokc.Item("PR").Item("OEM_NAME").Count
str1 = mokc.Item("PR").Item("OEM_NAME").Item(i).Key
If InStr(str1, "\") > 0 Then b_c = False
If InStr(str1, "/") > 0 Then b_c = False
If InStr(str1, "*") > 0 Then b_c = False
If InStr(str1, ":") > 0 Then b_c = False
If InStr(str1, "?") > 0 Then b_c = False
If b_c = False Then Exit For
Next
'N/A ֱ�� ��д N/A
'��Ӧ�����Ʋ��ܰ��������ַ�
'=====================
'���PR ģ���Ƿ���ڣ������ڿ���һ�ݵ���ǰĿ¼
If b_c = False Then
Msgbox "manufature name contain   \ / : * ?  please modify ��" & str1
Exit Sub
End If





'=================================
'���Ӽ�(TKSE) ������ͼ��
If mokc.Item("PR").Item("OEM_NAME").Count > 0 Then
If Not (mokc.Item("PR").Item("OEM_NAME").Item("TKSE") Is Nothing) Then
For i = 1 To mokc.Item("PR").Item("OEM_NAME").Item("TKSE").Count
str1 = mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item("TKSE").Item(i).Key)).Item("TKID").Key
If Len(str1) = 0 Then
Msgbox "manuture TKSY must have tkid ,Row��" & mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item("TKSE").Item(i).Key)).Key
b_c = False
End If
Next
End If
End If
'���Ӽ�(TKSE) ������ͼ��
'=================================
If b_c = False Then
Msgbox "Mechanical parts (TKSE) NO tkid can not create PR!"
Exit Sub
End If





'=================================
'��NA��N/A�����ͺź�TKID����ͬʱΪ��
For i = 1 To mokc.Item("WS_PartSingle").Item("WS_BODY").Count
str1 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Item("TKID").Key)
str2 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Item("OEM_ID").Key)
str3 = mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Item("OEM_NAME").Key
If (str3 <> "NA") Or (str3 <> "N/A") Then
If Len(str1) = 0 And Len(str2) = 0 Then
b_c = False
Msgbox "�Ǳ�����ͺţ���ɭ�Ų���ͬʱΪ�ա�����У�" & mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Key
Exit For
End If
End If
Next
'��NA��N/A�����ͺź�TKID����ͬʱΪ��
'=================================
If b_c = False Then
Msgbox "�����Ų���Ϊ��,�޷�����PR"
Exit Sub
End If



'20181021 PR���ͺ�����
sort_pr mokc

'20181021 PR���ͺ�����



'==============================
'�ڴ�������ȫ��PR����
For i = 1 To mokc.Item("PR").Item("OEM_NAME").Count
'�������PR��



'B
'SAP Item No.
'mokc.Item("PR").Item("OEM_NAME").Item(i).Add "B", "B"

For j = 1 To mokc.Item("PR").Item("OEM_NAME").Item(i).Count
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add CStr(j), "B"

'C
'Item ="PX00010001"
If i_moid = 0 Then
str1 = "PX" & mokc.Item("PR").Item("PRN_LAST").Key
Else
str1 = "MO" & mokc.Item("PR").Item("PRN_LAST").Key
End If

str2 = CStr(j)
'str2 = Left("000", 4 - Len(str2)) & str2
'20180530 �޸ĺ����ʽPX00010001ΪPX0001.001
str2 = Left(".00", 4 - Len(str2)) & str2

str1 = str1 & str2
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str1, "C"


'D
'ShortText, ���Ӽ� TKID���⹺�� OEM_ID
'1.���Ӽ�=��TKID
'2.�⹺����ͬʱ���ͺţ�����TKID,��������D���ͺţ�TKID���������ݺϲ���MEMO��
'3.�⹺���������ͺ�
'4.�⹺��������TKID
'5.�⹺����û�ͺ�Ҳû��TKID
str1 = mokc.Item("PR").Item("OEM_NAME").Item(i).Key
If str1 = "TKSE" Then
str2 = mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("TKID").Key
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str2, "D"
Else
str2 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("TKID").Key)
str3 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("OEM_ID").Key)
If Len(str3) > 0 Then
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str3, "D"
ElseIf Len(str2) > 0 Then
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str2, "D"
End If
End If



'E
'ֱ�ӽ�TKID_SUBASS����
str1 = mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("TKID_SUBASS").Key
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str1, "E"

'F
'ֱ�ӽ�OEM_NAME ����
str1 = mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("OEM_NAME").Key
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str1, "F"

'G
'����
'1.���Ӽ�.TKID**����
'2.�⹺��.����
str1 = mokc.Item("PR").Item("OEM_NAME").Item(i).Key
If str1 = "TKSE" Then
str2 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("TKID").Key)
str3 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("ItemName").Key)
If Len(str3) = 0 Then
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str2, "G"
Else
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str2 & "**" & str3, "G"
End If
Else
str3 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("ItemName").Key)
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str3, "G"
End If


'H
'CostUnit
'ʹ�� ���ٱ��������4λ CN.3  & ��λ������Ŀ�� & 41 & ��λ���ڹ�λ��
str1 = Left(wb_fl.Name, 4)
str2 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("TKID_STATION").Key)



'2019 12 06 ����ֱ������ɱ����ĺ�
If str2 Like "CN.??????.???*" Then
str3 = str2
Else
str2 = Left(str2, 11)
str3 = str1 & Mid(str2, 3, 5) & ".41" & Right(str2, 3)
End If
'2019 12 06 ����ֱ������ɱ����ĺ�


mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str3, "H"



'I
'����
str1 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("QTY").Key)
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str1, "I"


'J
'��λ
str1 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("UNIT").Key)
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str1, "J"


'L,M
'COST ELEMENT
'Other manufacturing material (Non-Independent Function) 40250000
's_str2 = "Other manufacturing material (Non-Independent Function)": s_str3 = "40250000": mokc.Item("PCE").Add s_str2, s_str2: mokc.Item("PCE").Item(s_str2).Add s_str3, s_str3
'Electrical  Parts Purchase  40270000
's_str2 = "Electrical  Parts Purchase": s_str3 = "40270000": mokc.Item("PCE").Add s_str2, s_str2: mokc.Item("PCE").Item(s_str2).Add s_str3, s_str3
'Pneumatic & Hydraulic   40280000
's_str2 = "Pneumatic & Hydraulic": s_str3 = "40280000": mokc.Item("PCE").Add s_str2, s_str2: mokc.Item("PCE").Item(s_str2).Add s_str3, s_str3
'Machinery & tooling (Single Part)   43202000
's_str2 = "Machinery & tooling (Single Part)": s_str3 = "43202000": mokc.Item("PCE").Add s_str2, s_str2: mokc.Item("PCE").Item(s_str2).Add s_str3, s_str3
If mokc.Item("WS_PartSingle").Item("M_C_P").Key = "Mechanics" Then
str1 = mokc.Item("PR").Item("OEM_NAME").Item(i).Key
If str1 = "TKSE" Then
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add "Machinery & tooling (Single Part)", "L"
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add "43202000", "M"
Else
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add "Other manufacturing material (Non-Independent Function)", "L"
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add "40250000", "M"
End If
ElseIf mokc.Item("WS_PartSingle").Item("M_C_P").Key = "Controls" Then
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add "Electrical  Parts Purchase", "L"
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add "40270000", "M"
ElseIf mokc.Item("WS_PartSingle").Item("M_C_P").Key = "Pneumatics" Then
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add "Pneumatic & Hydraulic", "L"
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add "40280000", "M"
End If





'N������
'�������ڸ�ʽת��
str1 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("E_DATE").Key)
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str1, "N"





'O��ע
'���������ע
'1.���Ӽ���TKSE�������**Description
'2.�⹺��, TKID**���**Description
str1 = mokc.Item("PR").Item("OEM_NAME").Item(i).Key
If str1 = "TKSE" Then
str2 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("DIMENSION").Key)
str3 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("DESC").Key)
If Len(str2) = 0 Or Len(str3) = 0 Then
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str2 & str3, "O"
Else
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str2 & "**" & str3, "O"
End If
Else
str2 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("TKID").Key)
str3 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("DIMENSION").Key)
If Len(str2) = 0 Or Len(str3) = 0 Then
str2 = str2 & str3
Else
str2 = str2 & "**" & str3
End If
str3 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("DESC").Key)
If Len(str2) = 0 Or Len(str3) = 0 Then
str2 = str2 & str3
Else
str2 = str2 & "**" & str3
End If
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str2, "O"
End If

If i_moid > 0 Then
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("O").Key = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("O").Key & "_MO" & Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("MO ID").Key)

End If





Next


'PR���ż�1
i_PRN_LAST = CInt(mokc.Item("PR").Item("PRN_LAST").Key)
i_PRN_LAST = i_PRN_LAST + 1
s_PRN_LAST = CStr(i_PRN_LAST)
If Len(s_PRN_LAST) < 4 Then
s_PRN_LAST = Left("000", 4 - Len(s_PRN_LAST)) & s_PRN_LAST
End If
mokc.Item("PR").Item("PRN_LAST").Key = s_PRN_LAST
'PR���ż�1

 

Next
'�ڴ�������ȫ��PR����
'==============================



'====================================
'��Ԫ�񳤶Ƚ�����35������Ķ� mokc.Item("PR").Item("OEM_NAME") �Ĵ���
'D�г��� ��..��������O��
'G�г��� ��##��������O��
'O�г��� ��^^��ǰ��ķֿ���������� ע��
For i = 1 To mokc.Item("PR").Item("OEM_NAME").Count
For j = 1 To mokc.Item("PR").Item("OEM_NAME").Item(i).Count
s_t1 = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("D").Key
s_t2 = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("G").Key
s_t3 = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("O").Key
s_t4 = ""

If Len(s_t1) > 35 Then
s_t4 = ".." & Right(s_t1, Len(s_t1) - 33)
s_t1 = Left(s_t1, 33) & ".."
End If
If Len(s_t2) > 35 Then
s_t4 = s_t4 & "##" & Right(s_t2, Len(s_t2) - 33)
s_t2 = Left(s_t2, 33) & "##"
End If
If Len(s_t4) > 0 And Len(s_t3) > 0 Then
s_t4 = s_t4 & "^^" & s_t3
Else
s_t4 = s_t4 & s_t3
End If
If Len(s_t4) <= 35 Then
s_t3 = s_t4
s_t4 = ""
Else
s_t3 = Left(s_t4, 35)
s_t4 = Right(s_t4, Len(s_t4) - 35)
End If
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("D").Key = s_t1
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("G").Key = s_t2
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("O").Key = s_t3
If Len(s_t4) > 0 Then
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add s_t4, "Comment"
End If
Next
Next
'��Ԫ�񳤶Ƚ�����35������Ķ� mokc.Item("PR").Item("OEM_NAME") �Ĵ���
'====================================








'==============================
'�ڴ���PR������������PR�ļ�
'��ģ��


For i = 1 To mokc.Item("PR").Item("OEM_NAME").Count
'PR���ļ�����PX1515_CN.305587-8-9_Spare parts_20170725.xlsm
'PR���ļ�����PX####_CN.######_OEM_NAME_YYYYMMDD.xlsm
str1 = Left(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(1).Item("C").Key, 6)


str2 = Left(wb_fl.Name, 9)
str3 = mokc.Item("PR").Item("OEM_NAME").Item(i).Key

'20181129 ȥ����Ӧ����������������ַ�
If Not mokc_manu.Item("MANUFATURE") Is Nothing Then
If Not mokc_manu.Item("MANUFATURE").Item("BODY").Item(str3) Is Nothing Then
str3 = mokc_manu.Item("MANUFATURE").Item("BODY").Item(str3).Item(str3).Item(1).Item(2).Key
End If
End If
'20181129 ȥ����Ӧ����������������ַ�


fln = str1 & "_" & str2 & "_" & str3 & "_" & Format(Now(), "YYYYMMDD") & ".xlsm"


open_wb wb_pr, mokc.Item("PR").Item("FLFP_TEMPLATE").Key


wb_pr.SaveAs mokc.Item("FL").Item("FDN").Item("FDNPR").Key & "\" & fln

wb_pr.Worksheets(1).Range("O7") = str1



'=======================================
'������Ԫ��Name of component .TK Internal Ident. number
wb_pr.Worksheets(1).Range("D20") = "Vendor Part No."
wb_pr.Worksheets(1).Range("G20") = "Name of component .TK Internal Ident. number"
'������Ԫ��Name of component .TK Internal Ident. number
'=======================================




'=========================Applicant:
 str1 = Application.UserName
If Len(str1) > 12 Then str1 = Environ("username")
If Len(str1) > 12 Then str1 = Left(str1, 12)
wb_pr.Worksheets(1).Range("C3") = str1
'=========================Applicant:


'=========================Application Date:
str1 = Format(Now(), "MM/DD/YYYY")
wb_pr.Worksheets(1).Range("M3") = str1
'=========================Application Date:




For j = 1 To mokc.Item("PR").Item("OEM_NAME").Item(i).Count

wb_pr.Worksheets(1).Range("B" & j + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("B").Key
wb_pr.Worksheets(1).Range("C" & j + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("C").Key
wb_pr.Worksheets(1).Range("D" & j + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("D").Key
wb_pr.Worksheets(1).Range("E" & j + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("E").Key
wb_pr.Worksheets(1).Range("F" & j + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("F").Key
wb_pr.Worksheets(1).Range("G" & j + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("G").Key
wb_pr.Worksheets(1).Range("H" & j + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("H").Key
wb_pr.Worksheets(1).Range("I" & j + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("I").Key
wb_pr.Worksheets(1).Range("J" & j + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("J").Key
'wb_pr.Worksheets(1).Range("K" & j + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("K").Key
wb_pr.Worksheets(1).Range("L" & j + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("L").Key
wb_pr.Worksheets(1).Range("M" & j + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("M").Key
wb_pr.Worksheets(1).Range("N" & j + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("N").Key
wb_pr.Worksheets(1).Range("O" & j + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("O").Key

If Not (mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("Comment") Is Nothing) Then
add_comm mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("Comment").Key, wb_pr.Worksheets(1), j + 20, 15, False
wb_pr.Worksheets(1).Rows(20 + j & ":" & 20 + j).Interior.Color = 255
End If

'���ܱ������¼PR��
ws_partsingle.Cells(CInt(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Key), CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("PA_Index").Item(1).Key)) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("C").Key

'���ܱ������¼PR��




Next


'���ô�ӡ����

wb_pr.Worksheets("PA").PageSetup.PrintArea = "$B$1:$O$" & j + 20

wb_pr.Save

'���ô�ӡ����
wb_pr.Close


Next

'�ڴ���PR������������PR�ļ�
''==============================


wb_fl.Save


'���ļ���
Shell "explorer.exe " & mokc.Item("FL").Item("FDN").Item("FDNPR").Key, vbNormalFocus
'���ļ���

Workbooks("PR_Create_Tool.xlsm").Saved = True
Workbooks("PR_Create_Tool.xlsm").Close

End Sub


Function TableHead_REC(mokc As OneKeyCls, ws As Worksheet, row_n As Integer, col_n As Integer, rg_value As String, skey As String) As Boolean
'��¼��ͷ
TableHead_REC = False
    If ws.Cells(row_n, col_n) = rg_value Then
        If mokc.Item(skey) Is Nothing Then
        mokc.Add skey, skey
        mokc.Item(skey).Add CStr(col_n)
        mokc.Item(skey).Add ""
        
        TableHead_REC = True
        End If
    Else
    'MsgBox "Can not find Table Head: cells(" & CStr(row_n) & "," & CStr(col_n) & ")=" & rg_value
    End If
End Function

Function OEM_NAME_check(str1 As String) As String
'��Ӧ��������ϴ
If str1 = "TKSY" Or str1 = "TK SY" Or str1 = "TK SE" Or str1 = "tkSY" Or str1 = "tk SY" Then
OEM_NAME_check = "TKSE"
ElseIf str1 Like "*��*��*" Then
Msgbox "��ȷ���Ƿ��Ѳɹ���" & str1 & Chr(10) & "�Ѳɹ���Ŀ ���Done ���������б���"
OEM_NAME_check = ""
ElseIf InStr(str1, Chr(10)) > 0 Or InStr(str1, Chr(13)) > 0 Or InStr(str1, "\") > 0 Or InStr(str1, "/") > 0 Then
Msgbox "Error Char in manufature name!"
OEM_NAME_check = ""
Else
OEM_NAME_check = str1
End If
'��Ӧ��������ϴ
End Function
Function UNIT_check(str1 As String) As String
'��λ������ϴ
'1."",�滻ΪPCE��Ϊ���ϴ���
'2."EA"�滻��ΪPCE
'3.ST �滻ΪSET
If str1 = "" Or str1 = "EA" Then
UNIT_check = "PCE"
ElseIf str1 = "LOT" Then
UNIT_check = "SET"
ElseIf str1 = "ST" Then
UNIT_check = "SET"
Else
UNIT_check = str1
End If


'��Ӧ��������ϴ
End Function

Function format_date_DDMMYYYY(m_c As Range) As String
'��ʽ�����ں���
'֧��Excelȫ�����ڸ�ʽ��CW������ʽ
    Dim date_1 As Date
    Dim s_1 As String
    Dim wk As Integer
    Dim str_date As String
'===============================
'��Ԫ���Ѿ������ڸ�ʽ�ģ����и�ʽת��

    If IsDate(m_c) = True Then
    date_1 = m_c
    format_date_DDMMYYYY = Format(date_1, "DD.MM.YYYY")
    Else
    format_date_DDMMYYYY = Trim(m_c)
    End If
'��Ԫ���Ѿ������ڸ�ʽ�ģ����и�ʽת��
'===============================
'===========================
'�ж��Ƿ�ת���ɹ������δ�ɹ����ж��Ƿ�ΪCW##��ʽ��ת��

If format_date_DDMMYYYY Like "##.##.####" Then
'�ɹ�ֱ������
Else

    str_date = format_date_DDMMYYYY
    If str_date Like "CW?/????" Then
    str_date = "CW0" & Right(str_date, 6)
    ElseIf str_date Like "CW?/????*" Then
    str_date = "CW0" & Mid(str_date, 3, 6)
    End If
    
    If str_date Like "CW??*" Then
    'Return the sunday of special week
    wk = CInt(Mid(str_date, 3, 2))
    Dim InputNum As Integer, FirstD As Date, StartD As Date, i As Integer
    InputNum = Val(wk)
    FirstD = CDate(Year(Date) & "-1" & "-1")
    StartD = FirstD + (InputNum - 1) * 7 - Weekday(FirstD, vbMonday) + 1
    date_1 = CDate(StartD + 4)
    If date_1 < Now() Then
    If str_date Like "CW??*" Then
    'Return the sunday of special week
    wk = CInt(Mid(str_date, 3, 2))
    InputNum = Val(wk)
    FirstD = CDate((Year(Date) + 1) & "-1" & "-1")
    StartD = FirstD + (InputNum - 1) * 7 - Weekday(FirstD, vbMonday) + 1
    date_1 = CDate(StartD + 4)
    ElseIf str_date Like "????-*-*" Then
    'Return the Change directly
    date_1 = CDate(str_date)
    End If
    End If
    format_date_DDMMYYYY = Format(date_1, "DD.MM.YYYY")
    End If
    
End If
'�ж��Ƿ�ת���ɹ������δ�ɹ����ж��Ƿ�ΪCW##��ʽ��ת��
'===========================


'If m_c.Comment Is Nothing Then m_c.AddComment
'm_c.Comment.Text Text:=CStr(m_c)
'm_c.NumberFormat = "yyyy/mm/dd;@"
'm_c = date_1






End Function


Sub PE_Creater()
Attribute PE_Creater.VB_ProcData.VB_Invoke_Func = "e\n14"
'�������ڴ������� PR����
'MsgBox "WinShuttle Can Use again! Please UPload to SAP after Create PR!"
 back_followinglist
 If Enagble_addins("PPPE") Then
 End If

Dim wb_pr As Workbook

Dim mfso As New CFSO
Dim mokc As New OneKeyCls
mokc.Add "FL", "FL"
mokc.Add "PR", "PR"
mokc.Item("PR").Add "FLFP_TEMPLATE", "FLFP_TEMPLATE"
mokc.Item("PR").Add "OEM_NAME", "OEM_NAME"
mokc.Add "WS_PartSingle", "WS_PartSingle"
mokc.Item("WS_PartSingle").Add "WS_HEAD", "WS_HEAD"
mokc.Item("WS_PartSingle").Add "WS_BODY", "WS_BODY"


mokc.Item("FL").Add "FLN", "FLN"
mokc.Item("FL").Add "FDN", "FDN"

mokc.Item("PR").Add "PRN_LAST", "PRN_LAST"

mokc.Item("FL").Item("FDN").Add "FDNPR", "FDNPR"

mokc.Item("FL").Add "CUR_PR_NUM", "CUR_PR_NUM"


'=======================
'�ҵ��Ѿ��򿪵Ļ�е���ٱ�
Dim b_c As Boolean
Dim str1 As String
Dim str2 As String
Dim str3 As String
Dim str4 As String

Dim s_t1 As String
Dim s_t2 As String
Dim s_t3 As String
Dim s_t4 As String


Dim wb_fl As Workbook
Dim i As Long
Dim i_last As Long
Dim i_PRN_LAST As Integer
Dim s_PRN_LAST As String
Dim fln As String
Dim fdn As String
Dim ws_partsingle As Worksheet

Dim j As Long
Dim j_last As Long

'��ȡPartSingle����ĵ�Ԫ��
Dim POS As String
Dim QTY As String
Dim UNIT As String
Dim ItemName As String
Dim OEM_ID As String
Dim OEM_NAME As String
Dim TKID_SUBASS As String
Dim TKID_STATION As String
Dim PA_Index As String
Dim R_DATE As String
Dim E_DATE As String
Dim MATERIAL As String
Dim STANDARD As String
Dim DIMENSION As String
Dim dbl_qty As Double



'��ȡPartSingle����ĵ�Ԫ��




b_c = False
For i = 1 To Workbooks.Count
str1 = Workbooks(i).Name
If str1 Like "CN.*ontrols*" Then
Set wb_fl = Workbooks(i)
b_c = True
Exit For
End If
Next
If Not (wb_fl Is Nothing) Then
Msgbox "����PR�ĸ��ٱ�Ϊ��" & Chr(10) & wb_fl.Name
Msgbox "Create PR For:" & Chr(10) & wb_fl.Name


If wb_fl.ReadOnly = True Then
Msgbox "ֻ����ʽ�ĸ��ٱ��޷�����PR"
Msgbox "Should Not be Read Only."

wb_fl.Close
Exit Sub
End If
Else
Msgbox "���ȴ򿪸��ٱ� CN.*ontrols*Following.xlsm "
Msgbox "Please Open: CN.*Controls*Following.xlsm "

Exit Sub
End If
'�ҵ��Ѿ��򿪵Ļ�е���ٱ�
'=======================
mokc.Item("FL").Item("FLN").Key = wb_fl.Name
mokc.Item("FL").Item("FDN").Key = wb_fl.Path
If Right(mokc.Item("FL").Item("FDN").Key, 1) <> "\" Then
mokc.Item("FL").Item("FDN").Key = mokc.Item("FL").Item("FDN").Key & "\"
End If
'=======================
'Checking  �ļ����Ƿ���� (mokc.Item("FL").Item("FDN").Item("FDNPR").Key)
fdn = mokc.Item("FL").Item("FDN").Key & "PAE_Controls&Robot&Commissioning"
mokc.Item("FL").Item("FDN").Item("FDNPR").Key = fdn
If mfso.folderexists(fdn) = False Then
If Msgbox("Folder Not exist" & Chr(10) & fdn & Chr(10) & "Create Press: OK ", vbOKCancel) = vbOK Then
mfso.CreateFolder fdn
Msgbox "Already Create:" & Chr(10) & mokc.Item("FL").Item("FDN").Key & "PAE_Controls&Robot&Commissioning" & Chr(10) & "Def PR Number from 0002 "
Else
Exit Sub
End If
End If
'Checking  �ļ����Ƿ���� (mokc.Item("FL").Item("FDN").Item("FDNPR").Key)
'=======================







'=======================
'Get_Last PAE Number

i_PRN_LAST = 1
s_PRN_LAST = "0001"
Record_file_in_folder mokc.Item("FL").Item("FDN").Item("FDNPR"), mokc.Item("FL").Item("FDN").Item("FDNPR").Key, ".xlsm"
For i = 1 To mokc.Item("FL").Item("FDN").Item("FDNPR").Item("FILE").Count
fln = mokc.Item("FL").Item("FDN").Item("FDNPR").Item("FILE").Item(i).Item("FLN").Key
fdn = mokc.Item("FL").Item("FDN").Item("FDNPR").Item("FILE").Item(i).Item("FDN").Key
If fdn = mokc.Item("FL").Item("FDN").Item("FDNPR").Key Then
    If fln Like "P?####*.xlsm" Then
    str1 = Mid(fln, 3, 4)
        If CInt(str1) > CInt(s_PRN_LAST) Then
        s_PRN_LAST = str1
        i_PRN_LAST = CInt(s_PRN_LAST)
       
        
        
        End If
    ElseIf fln Like "PAE###*.xlsm" Then
    
    
        str1 = Mid(fln, 4, 3)
        If CInt(str1) > CInt(s_PRN_LAST) Then
        s_PRN_LAST = str1
        i_PRN_LAST = CInt(s_PRN_LAST)
       
        
        End If
    Else
    
    Msgbox fln
    End If
End If
Next
i_PRN_LAST = i_PRN_LAST + 1
s_PRN_LAST = CStr(i_PRN_LAST)
If Len(s_PRN_LAST) < 4 Then
s_PRN_LAST = Left("000", 4 - Len(s_PRN_LAST)) & s_PRN_LAST
End If
mokc.Item("PR").Item("PRN_LAST").Key = s_PRN_LAST
'Get_Last PAE Number
'=======================




'=====================
'��ȡ Part_single ���
b_c = True ' b_c = True  Part_Single ��ȡ�ɹ���b_c = False  Part_Single ��ȡʧ��
mokc.Item("WS_PartSingle").Add "M_C_P", "M_C_P"
'Record the pype: Controls or Mechanics or Pneumatics
'������ Parts_Single,ʧ��
If b_c Then
If ws_exist(wb_fl, "Parts_Single") = False Then
b_c = False
Msgbox "Following list  not exist Parts_Single"
Else
Set ws_partsingle = wb_fl.Worksheets("Parts_Single")
End If
End If
'�޷��ж��� �� or �� or �� ʧ��
If b_c Then
If InStr(wb_fl.Name, "Mechanics") > 0 Then
mokc.Item("WS_PartSingle").Item("M_C_P").Key = "Mechanics"
ElseIf InStr(wb_fl.Name, "Controls") > 0 Then
mokc.Item("WS_PartSingle").Item("M_C_P").Key = "Controls"
ElseIf InStr(wb_fl.Name, "Pneumatics") > 0 Then
mokc.Item("WS_PartSingle").Item("M_C_P").Key = "Pneumatics"
Else
b_c = False
Msgbox wb_fl.Name & Chr(10) & "File name must contain Mechanics  or Controls orPneumatics"
End If
End If
'�ж� Part_Single ��ʽ�Ƿ���Ԥ�ø�ʽ
If b_c Then
If mokc.Item("WS_PartSingle").Item("M_C_P").Key = "Controls" Then

'Template 1
'   If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 1, "Pos.", "POS") Then b_c = False
'    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 3, "Qty", "QTY") Then b_c = False
'    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 4, "Base Unit", "UNIT") Then b_c = False
'    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 6, "Matl. Descrip.", "ItemName") Then b_c = False
'    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 7, "Material No.", "TKID") Then b_c = False
'    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 8, "SPI", "SPI") Then b_c = False
'    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 11, "Manuf.Part.No.", "OEM_ID") Then b_c = False
'    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 12, "Basic Material", "MATERIAL") Then b_c = False
'    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 13, "Matl. Standard", "STANDARD") Then b_c = False
'    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 14, "Size/Dimension", "DIMENSION") Then b_c = False
'    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 15, "Manufacturer", "OEM_NAME") Then b_c = False
'    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 16, "Sub-Assy._Number", "TKID_SUBASS") Then b_c = False
'    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 17, "Station_Number", "TKID_STATION") Then b_c = False
'    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 18, "PA_Index", "PA_Index") Then If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 18, "PA_Number", "PA_Index") Then b_c = False
'    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 19, "Release_date", "R_DATE") Then b_c = False
'    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 20, "Expect Week", "E_DATE") Then b_c = False
'    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 21, "description", "DESC") Then If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 21, "Description", "DESC") Then b_c = False
'    If b_c = False Then MsgBox "Part_Single ��ͷ�޷�ʶ��"
    
    
    
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 7, 2, "Pos #", "POS") Then b_c = False
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 7, 4, "Qty", "QTY") Then b_c = False
    
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 7, 7, "Q" & Chr(10) & "C", "UNIT") Then b_c = False
    
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 7, 5, "Description", "ItemName") Then b_c = False
    
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 7, 3, "TK Ident Number", "TKID") Then b_c = False
    
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 7, 8, "SW", "SPI") Then b_c = False
    
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 7, 10, "Manufacturer Part Number", "OEM_ID") Then b_c = False
    
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 7, 9, "Material Name", "MATERIAL") Then b_c = False
    
    'If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 13, "Matl. Standard", "STANDARD") Then b_c = False
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 7, 6, " TECHNICAL PARAMETERS", "DIMENSION") Then b_c = False
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 7, 11, "Manufacturer", "OEM_NAME") Then b_c = False
    
    'If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 16, "Sub-Assy._Number", "TKID_SUBASS") Then b_c = False
    
    'If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 17, "Station_Number", "TKID_STATION") Then b_c = False
    
    '�����ɱ����ĺ���
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 7, 15, "Cost Unit", "WBS") Then b_c = False
    '�����ɱ����ĺ���
    
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 7, 16, "PA_Index", "PA_Index") Then b_c = False
    
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 7, 17, "Release_date", "R_DATE") Then b_c = False
    
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 7, 18, "Expect " & Chr(10) & "Week", "E_DATE") Then b_c = False
    
    
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 7, 20, "Description", "DESC") Then b_c = False
    If b_c = False Then Msgbox "Part_Single ��ͷ�޷�ʶ��"
    
    
End If
End If



If b_c Then

i_last = ws_partsingle.UsedRange.Rows(ws_partsingle.UsedRange.Rows.Count).row
For i = 8 To i_last
'��ȡ������Ϊ�㣬PA_Index Ϊ�յ���




PA_Index = Trim(ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("PA_Index").Item(1).Key)))
QTY = Trim(ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("QTY").Item(1).Key)))


If Len(PA_Index) = 0 And Len(QTY) > 0 Then
OEM_NAME = Trim(ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("OEM_NAME").Item(1).Key)))
'TKID_STATION = Trim(ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("TKID_STATION").Item(1).Key)))
dbl_qty = 0
Str_TO_Dbl QTY, dbl_qty
If Not (dbl_qty > 0) Then
Msgbox "Qty Must >0  check Row:" & i
b_c = False
End If
If b_c Then
If Len(OEM_NAME) = 0 Then
Msgbox "Manufature can not be empty ,ROW:" & i
b_c = False
End If
If b_c Then
mokc.Item("WS_PartSingle").Item("WS_BODY").Add CStr(i), CStr(i)
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("POS").Item(1).Key)), "POS"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("QTY").Item(1).Key)), "QTY"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("UNIT").Item(1).Key)), "UNIT"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("ItemName").Item(1).Key)), "ItemName"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("TKID").Item(1).Key)), "TKID"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("SPI").Item(1).Key)), "SPI"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("OEM_ID").Item(1).Key)), "OEM_ID"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("MATERIAL").Item(1).Key)), "MATERIAL"
'mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("STANDARD").Item(1).Key)), "STANDARD"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("DIMENSION").Item(1).Key)), "DIMENSION"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("OEM_NAME").Item(1).Key)), "OEM_NAME"
'mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("TKID_SUBASS").Item(1).Key)), "TKID_SUBASS"
'mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("TKID_STATION").Item(1).Key)), "TKID_STATION"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("PA_Index").Item(1).Key)), "PA_Index"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("R_DATE").Item(1).Key)), "R_DATE"
'add WBS
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("WBS").Item(1).Key)), "WBS"
'add WBS
E_DATE = format_date_DDMMYYYY(ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("E_DATE").Item(1).Key)))
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add E_DATE, "E_DATE"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("DESC").Item(1).Key)), "DESC"

End If
End If
End If
Next
End If
'��ȡ Part_single ���
'=====================

If b_c = False Then
Msgbox "��ȡ Part_Single ʧ�ܣ��޷�����PR"
Msgbox "Reading Part_Single failure can not create PR"
Exit Sub
End If





'=====================
'���PR ģ���Ƿ���ڣ������ڿ���һ�ݵ���ǰĿ¼
'Z:\24_Temp\PA_Logs\V1.1\TEMPLATE\010c1612_Purchase Requisition.xlsm
If mfso.FileExists(mokc.Item("FL").Item("FDN").Key & "\010c1612_Purchase Requisition.xlsm") Then
Else
 If mfso.FileExists("Z:\24_Temp\PA_Logs\V1.1\TEMPLATE\010c1612_Purchase Requisition.xlsm") Then
 mfso.copy_file "Z:\24_Temp\PA_Logs\V1.1\TEMPLATE\010c1612_Purchase Requisition.xlsm", mokc.Item("FL").Item("FDN").Key & "\010c1612_Purchase Requisition.xlsm"
 Else
 Msgbox "�� PRģ�壺Z:\24_Temp\PA_Logs\V1.1\TEMPLATE\010c1612_Purchase Requisition.xlsm"
 Msgbox "NO pr template : Z:\24_Temp\PA_Logs\V1.1\TEMPLATE\010c1612_Purchase Requisition.xlsm"
  
 b_c = False
 End If
End If
mokc.Item("PR").Item("FLFP_TEMPLATE").Key = mokc.Item("FL").Item("FDN").Key & "\010c1612_Purchase Requisition.xlsm"
'���PR ģ���Ƿ���ڣ������ڿ���һ�ݵ���ǰĿ¼
'=====================
If b_c = False Then
Msgbox "PR ģ�岻���ڣ��޷�����PR"
Msgbox "PR template not exist can not create pr"
Exit Sub
End If






'==============================
'���ģ������Ŀ����Ŀ���Ƿ���ڣ�������Ҫ������
If open_wb(wb_pr, mokc.Item("PR").Item("FLFP_TEMPLATE").Key) Then
str1 = wb_pr.Worksheets(1).Range("G7")
Do While Len(str1) = 0
str1 = InputBox("Project Number", "Template Create", "CN.#######")
If str1 = "CN.#######" Then str1 = ""
Loop
wb_pr.Worksheets(1).Range("G7") = str1
str1 = wb_pr.Worksheets(1).Range("M7")
Do While Len(str1) = 0
str1 = InputBox("Project Name", "Template Create")
Loop
wb_pr.Worksheets(1).Range("M7") = str1
wb_pr.Save
wb_pr.Saved = True
wb_pr.Close
Else
b_c = False
End If
'���ģ������Ŀ����Ŀ���Ƿ���ڣ�������Ҫ������
'==============================
If b_c = False Then
Msgbox "PR ģ�� ����Ŀ���ƣ�����Ŀ��δ��д"
Msgbox "NO Project name & Number in PR template."
Exit Sub
End If







'============================
'��λ ���
For i = 1 To mokc.Item("WS_PartSingle").Item("WS_BODY").Count
str1 = UNIT_check(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Item("UNIT").Key)
If Len(str1) = 0 Then
b_c = False
Exit For
Else
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Item("UNIT").Key = str1
End If
Next

'��λ ���
'============================
If b_c = False Then
Msgbox "Unit check error:" & Chr(10) & "Row:" & mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Key
Exit Sub
End If




'============================
'�������� ���
For i = 1 To mokc.Item("WS_PartSingle").Item("WS_BODY").Count
str1 = mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Item("E_DATE").Key
If Len(str1) = 0 Then
b_c = False
Exit For
Else
End If
Next

'�������� ���
'============================
If b_c = False Then
Msgbox "�������ڼ�� ʧ�� " & Chr(10) & "�к�:" & mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Key

Msgbox "wrong date " & Chr(10) & "�к�:" & mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Key
Exit Sub
End If






'============================
'COST_UNIT ���
For i = 1 To mokc.Item("WS_PartSingle").Item("WS_BODY").Count
str1 = COST_UNIT_check(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Item("WBS").Key, mokc.Item("FL").Item("FLN").Key)
If Len(str1) = 0 Then
b_c = False
Exit For
Else
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Item("WBS").Key = str1
End If
Next

'COST_UNIT ���
'============================
If b_c = False Then
Msgbox "Cost unit check error:" & Chr(10) & "row:" & mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Key
Exit Sub
End If






'============================
'��Ӧ������ ��鼰����
For i = 1 To mokc.Item("WS_PartSingle").Item("WS_BODY").Count
str1 = OEM_NAME_check(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Item("OEM_NAME").Key)
If Len(str1) = 0 Then
b_c = False
Exit For
End If
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Item("OEM_NAME").Key = str1


If mokc.Item("PR").Item("OEM_NAME").Item(str1) Is Nothing Then
mokc.Item("PR").Item("OEM_NAME").Add str1, str1
mokc.Item("PR").Item("OEM_NAME").Item(str1).Add CStr(i), CStr(i)


Else
mokc.Item("PR").Item("OEM_NAME").Item(str1).Add CStr(i), CStr(i)
End If
Next
'��Ӧ������ ��鼰����
'============================
If b_c = False Then
Msgbox "Manufature error:" & Chr(10) & "Row:" & mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Key
Exit Sub
End If


'============================
'��׼�� �� ���� PA ����ֿ���
If mokc.Item("PR").Item("OEM_NAME").Count > 1 Then
If Not (mokc.Item("PR").Item("OEM_NAME").Item("NA") Is Nothing) Then
Msgbox "Row ��" & mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item("NA").Item(1).Key)).Key
b_c = False
End If
End If
'��׼�� �� ���� PA ����ֿ���
'============================
If b_c = False Then
Msgbox "Manufacturer cannot be N/A"
Exit Sub
End If



'=====================
'N/A ֱ�� ��д N/A
'��Ӧ�����Ʋ��ܰ��������ַ�
If mokc.Item("PR").Item("OEM_NAME").Count > 1 Then
If Not (mokc.Item("PR").Item("OEM_NAME").Item("N/A") Is Nothing) Then
For i = 1 To mokc.Item("PR").Item("OEM_NAME").Item("N/A").Count
ws_partsingle.Cells(CInt(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item("N/A").Item(i).Key)).Key), CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("PA_Index").Item(1).Key)) = "N/A"
Next
End If
End If

If Not (mokc.Item("PR").Item("OEM_NAME").Item("N/A") Is Nothing) Then mokc.Item("PR").Item("OEM_NAME").Remove "N/A"

For i = 1 To mokc.Item("PR").Item("OEM_NAME").Count
str1 = mokc.Item("PR").Item("OEM_NAME").Item(i).Key
If InStr(str1, "\") > 0 Then b_c = False
If InStr(str1, "/") > 0 Then b_c = False
If InStr(str1, "*") > 0 Then b_c = False
If InStr(str1, ":") > 0 Then b_c = False
If InStr(str1, "?") > 0 Then b_c = False
If b_c = False Then Exit For
Next
'N/A ֱ�� ��д N/A
'��Ӧ�����Ʋ��ܰ��������ַ�
'=====================
'���PR ģ���Ƿ���ڣ������ڿ���һ�ݵ���ǰĿ¼
If b_c = False Then
Msgbox "��Ӧ�����ư��������ַ�  \ / : * ? �޷�����PR�����޸ģ�" & str1
Msgbox "Special characters  \ / : * ? Can not Create PR��" & str1
Exit Sub
End If





'=================================
'���Ӽ�(TKSE) ������ͼ��
If mokc.Item("PR").Item("OEM_NAME").Count > 0 Then
If Not (mokc.Item("PR").Item("OEM_NAME").Item("TKSE") Is Nothing) Then
For i = 1 To mokc.Item("PR").Item("OEM_NAME").Item("TKSE").Count
str1 = mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item("TKSE").Item(i).Key)).Item("TKID").Key
If Len(str1) = 0 Then
Msgbox "���Ӽ�������ͼ��,���ٱ��кţ�" & mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item("TKSE").Item(i).Key)).Key
Msgbox "TKSE MUST have TKID:" & mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item("TKSE").Item(i).Key)).Key

b_c = False
End If
Next
End If
End If
'���Ӽ�(TKSE) ������ͼ��
'=================================
If b_c = False Then
Msgbox "���Ӽ�(TKSE)�޵�ɭͼ��,�޷�����PR"
Msgbox "TKSE MUST have TKID"
Exit Sub
End If



'=================================
'��NA��N/A�����ͺź�TKID����ͬʱΪ��
For i = 1 To mokc.Item("WS_PartSingle").Item("WS_BODY").Count
str1 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Item("TKID").Key)
str2 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Item("OEM_ID").Key)
str3 = mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Item("OEM_NAME").Key
If (str3 <> "NA") Or (str3 <> "N/A") Then
If Len(str1) = 0 And Len(str2) = 0 Then
b_c = False
Msgbox "Short text Error: check Row:" & mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Key
Exit For
End If
End If
Next
'��NA��N/A�����ͺź�TKID����ͬʱΪ��
'=================================
If b_c = False Then
Msgbox "�����Ų���Ϊ��,�޷�����PR"

Msgbox "MUST have ShortText"
Exit Sub
End If


'20181021 ���ͺ�����

sort_pr mokc
'20181021 ���ͺ�����



'==============================
'�ڴ�������ȫ��PR����
For i = 1 To mokc.Item("PR").Item("OEM_NAME").Count
'�������PR��



'B
'SAP Item No.
'mokc.Item("PR").Item("OEM_NAME").Item(i).Add "B", "B"

For j = 1 To mokc.Item("PR").Item("OEM_NAME").Item(i).Count
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add CStr(j), "B"

'C
'Item ="PE00010001"
str1 = "PE" & mokc.Item("PR").Item("PRN_LAST").Key
str2 = CStr(j)

'str2 = Left("000", 4 - Len(str2)) & str2
'20180530 �޸ĺ����ʽPE00010001ΪPE0001.001
str2 = Left(".00", 4 - Len(str2)) & str2

str1 = str1 & str2
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str1, "C"


'D
'ShortText, ���Ӽ� TKID���⹺�� OEM_ID
'1.���Ӽ�=��TKID
'2.�⹺����ͬʱ���ͺţ�����TKID,��������D���ͺţ�TKID���������ݺϲ���MEMO��
'3.�⹺���������ͺ�
'4.�⹺��������TKID
'5.�⹺����û�ͺ�Ҳû��TKID
str1 = mokc.Item("PR").Item("OEM_NAME").Item(i).Key
If str1 = "TKSE" Then
str2 = mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("TKID").Key
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str2, "D"
Else
str2 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("TKID").Key)
str3 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("OEM_ID").Key)
If Len(str3) > 0 Then
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str3, "D"
ElseIf Len(str2) > 0 Then
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str2, "D"
End If
End If



'E
'�����Ĵ���Ϊ��
'str1 = mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("TKID_SUBASS").Key
'mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str1, "E"

'F
'ֱ�ӽ�OEM_NAME ����
str1 = mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("OEM_NAME").Key
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str1, "F"

'G
'����
'1.���Ӽ�.TKID**����
'2.�⹺�� �� ��쿹�ͨ����д ��������

str1 = mokc.Item("PR").Item("OEM_NAME").Item(i).Key
If str1 = "TKSE" Then
str2 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("TKID").Key)
str3 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("ItemName").Key)
If Len(str3) = 0 Then
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str2, "G"
Else
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str2 & "**" & str3, "G"
End If
Else
str3 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("DIMENSION").Key)
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str3, "G"
End If


'H
'CostUnit
'ʹ�� ���ٱ��������4λ CN.3  & ��λ������Ŀ�� & 41 & ��λ���ڹ�λ��

str3 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("WBS").Key)
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str3, "H"



'I
'����
str1 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("QTY").Key)
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str1, "I"


'J
'��λ
str1 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("UNIT").Key)
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str1, "J"


'L,M
'COST ELEMENT
'Other manufacturing material (Non-Independent Function) 40250000
's_str2 = "Other manufacturing material (Non-Independent Function)": s_str3 = "40250000": mokc.Item("PCE").Add s_str2, s_str2: mokc.Item("PCE").Item(s_str2).Add s_str3, s_str3
'Electrical  Parts Purchase  40270000
's_str2 = "Electrical  Parts Purchase": s_str3 = "40270000": mokc.Item("PCE").Add s_str2, s_str2: mokc.Item("PCE").Item(s_str2).Add s_str3, s_str3
'Pneumatic & Hydraulic   40280000
's_str2 = "Pneumatic & Hydraulic": s_str3 = "40280000": mokc.Item("PCE").Add s_str2, s_str2: mokc.Item("PCE").Item(s_str2).Add s_str3, s_str3
'Machinery & tooling (Single Part)   43202000
's_str2 = "Machinery & tooling (Single Part)": s_str3 = "43202000": mokc.Item("PCE").Add s_str2, s_str2: mokc.Item("PCE").Item(s_str2).Add s_str3, s_str3
If mokc.Item("WS_PartSingle").Item("M_C_P").Key = "Mechanics" Then
str1 = mokc.Item("PR").Item("OEM_NAME").Item(i).Key
If str1 = "TKSE" Then
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add "Machinery & tooling (Single Part)", "L"
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add "43202000", "M"
Else
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add "Other manufacturing material (Non-Independent Function)", "L"
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add "40250000", "M"
End If
ElseIf mokc.Item("WS_PartSingle").Item("M_C_P").Key = "Controls" Then
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add "Electrical  Parts Purchase", "L"
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add "40270000", "M"
ElseIf mokc.Item("WS_PartSingle").Item("M_C_P").Key = "Pneumatics" Then
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add "Pneumatic & Hydraulic", "L"
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add "40280000", "M"
End If





'N������
'�������ڸ�ʽת��
str1 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("E_DATE").Key)
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str1, "N"





'O��ע
'E�кϲ�I�кϲ�T��


str1 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("ItemName").Key) '5
str2 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("MATERIAL").Key) '9
str3 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("DESC").Key)
If InStr(str1, str2) > 0 Then
str1 = str1
ElseIf InStr(str2, str1) > 0 Then
str1 = str2
Else
str1 = str1 & str2
End If
If Len(str3) > 0 Then str1 = str1 & str3
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str1, "O"


'��� ����������Ϊ�գ�Ҳ����NAME��Ϊ�գ��򽫱�ע��Ų�뼼��������
str1 = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("G").Key
If Len(str1) = 0 Then
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("G").Key = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("O").Key
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("O").Key = ""
End If
'��� ����������Ϊ�գ�Ҳ����NAME��Ϊ�գ��򽫱�ע��Ų�뼼��������


Next






'PR���ż�1
i_PRN_LAST = CInt(mokc.Item("PR").Item("PRN_LAST").Key)
i_PRN_LAST = i_PRN_LAST + 1
s_PRN_LAST = CStr(i_PRN_LAST)
If Len(s_PRN_LAST) < 4 Then
s_PRN_LAST = Left("000", 4 - Len(s_PRN_LAST)) & s_PRN_LAST
End If
mokc.Item("PR").Item("PRN_LAST").Key = s_PRN_LAST
'PR���ż�1

 

Next
'�ڴ�������ȫ��PR����
'==============================


'====================================
'��Ԫ�񳤶Ƚ�����35������Ķ� mokc.Item("PR").Item("OEM_NAME") �Ĵ���
'D�г��� ��..��������O��
'G�г��� ��##��������O��
'O�г��� ��^^��ǰ��ķֿ���������� ע��
For i = 1 To mokc.Item("PR").Item("OEM_NAME").Count
For j = 1 To mokc.Item("PR").Item("OEM_NAME").Item(i).Count
s_t1 = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("D").Key
s_t2 = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("G").Key
s_t3 = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("O").Key
s_t4 = ""
If Len(s_t1) > 35 Then
s_t4 = ".." & Right(s_t1, Len(s_t1) - 33)
s_t1 = Left(s_t1, 33) & ".."
End If
If Len(s_t2) > 35 Then
s_t4 = s_t4 & "##" & Right(s_t2, Len(s_t2) - 33)
s_t2 = Left(s_t2, 33) & "##"
End If
If Len(s_t4) > 0 And Len(s_t3) > 0 Then
s_t4 = s_t4 & "^^" & s_t3
Else
s_t4 = s_t4 & s_t3
End If
If Len(s_t4) <= 35 Then
s_t3 = s_t4
s_t4 = ""
Else
s_t3 = Left(s_t4, 35)
s_t4 = Right(s_t4, Len(s_t4) - 35)
End If
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("D").Key = s_t1
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("G").Key = s_t2
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("O").Key = s_t3
If Len(s_t4) > 0 Then
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add s_t4, "Comment"
End If
Next
Next
'��Ԫ�񳤶Ƚ�����35������Ķ� mokc.Item("PR").Item("OEM_NAME") �Ĵ���
'====================================








'==============================
'�ڴ���PR������������PR�ļ�
'��ģ��


For i = 1 To mokc.Item("PR").Item("OEM_NAME").Count
'PR���ļ�����PX1515_CN.305587-8-9_Spare parts_20170725.xlsm
'PR���ļ�����PX####_CN.######_OEM_NAME_YYYYMMDD.xlsm
str1 = Left(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(1).Item("C").Key, 6)


str2 = Left(wb_fl.Name, 9)
str3 = mokc.Item("PR").Item("OEM_NAME").Item(i).Key
fln = str1 & "_" & str2 & "_" & str3 & "_" & Format(Now(), "YYYYMMDD") & ".xlsm"


open_wb wb_pr, mokc.Item("PR").Item("FLFP_TEMPLATE").Key
wb_pr.SaveAs mokc.Item("FL").Item("FDN").Item("FDNPR").Key & "\" & fln


'=======================================
'������Ԫ��Name of component .TK Internal Ident. number
wb_pr.Worksheets(1).Range("D20") = "Vendor Part No."
wb_pr.Worksheets(1).Range("G20") = "Name of component .TK Internal Ident. number"
'������Ԫ��Name of component .TK Internal Ident. number
'=======================================


wb_pr.Worksheets(1).Range("O7") = str1



'=========================Applicant:
 str1 = Application.UserName
If Len(str1) > 12 Then str1 = Environ("username")
If Len(str1) > 12 Then str1 = Left(str1, 12)
wb_pr.Worksheets(1).Range("C3") = str1
'=========================Applicant:


'=========================Application Date:
str1 = Format(Now(), "MM/DD/YYYY")
wb_pr.Worksheets(1).Range("M3") = str1
'=========================Application Date:



'��ͬWBS��SHORTTEXT ��Ҫ�ϲ���������ȡ�Ĳ����ǣ������ļ���д�����һ���� ����Ƿ���Ժϲ������Ժϲ���ϲ�
'i_PRN_LAST s_PRN_LAST ʵʱ����PR����д���ݣ�����ԭʼֵ
i_PRN_LAST = 1
s_PRN_LAST = "1"
Dim i_curr As Integer

Dim WBS_s As String
Dim SortText_s As String
Dim Memo_s As String



For j = 1 To mokc.Item("PR").Item("OEM_NAME").Item(i).Count


'===
WBS_s = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("H").Key
SortText_s = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("D").Key
Memo_s = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("O").Key



OEM_NAME = mokc.Item("PR").Item("OEM_NAME").Item(i).Key
If mokc.Item("MERGE") Is Nothing Then mokc.Add "MERGE", "MERGE"
If mokc.Item("MERGE").Item(OEM_NAME) Is Nothing Then mokc.Item("MERGE").Add OEM_NAME, OEM_NAME

If mokc.Item("MERGE").Item(OEM_NAME).Item(WBS_s & SortText_s & Memo_s) Is Nothing Then

b_c = True
i_curr = i_PRN_LAST
i_PRN_LAST = i_PRN_LAST + 1


mokc.Item("MERGE").Item(OEM_NAME).Add WBS_s & SortText_s & Memo_s, WBS_s & SortText_s & Memo_s
mokc.Item("MERGE").Item(OEM_NAME).Item(WBS_s & SortText_s & Memo_s).Add CStr(i_curr), "PRN"

QTY = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("I").Key
mokc.Item("MERGE").Item(OEM_NAME).Item(WBS_s & SortText_s & Memo_s).Add QTY, "QTY"


str1 = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("C").Key
str2 = CStr(i_curr)
str2 = Left(".000", 4 - Len(str2)) & str2
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("C").Key = Left(str1, 6) & str2





Else
b_c = False
i_curr = CInt(mokc.Item("MERGE").Item(OEM_NAME).Item(WBS_s & SortText_s & Memo_s).Item("PRN").Key)
QTY = CStr(CInt(mokc.Item("MERGE").Item(OEM_NAME).Item(WBS_s & SortText_s & Memo_s).Item("QTY").Key) + CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("I").Key))
mokc.Item("MERGE").Item(OEM_NAME).Item(WBS_s & SortText_s & Memo_s).Item("QTY").Key = QTY


str1 = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("C").Key
str2 = CStr(i_curr)
str2 = Left(".000", 4 - Len(str2)) & str2
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("C").Key = Left(str1, 6) & str2




b_c = True
End If



'===

If b_c Then
           'wb_pr.Worksheets(1).Range("B" & i_curr + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("B").Key
            wb_pr.Worksheets(1).Range("B" & i_curr + 20) = i_curr
            
            wb_pr.Worksheets(1).Range("C" & i_curr + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("C").Key
            wb_pr.Worksheets(1).Range("D" & i_curr + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("D").Key
            'wb_pr.Worksheets(1).Range("E" & j + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("E").Key
            wb_pr.Worksheets(1).Range("F" & i_curr + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("F").Key
            wb_pr.Worksheets(1).Range("G" & i_curr + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("G").Key
            wb_pr.Worksheets(1).Range("H" & i_curr + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("H").Key
            
            'wb_pr.Worksheets(1).Range("I" & j + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("I").Key
            
            wb_pr.Worksheets(1).Range("I" & i_curr + 20) = QTY
            
            
            wb_pr.Worksheets(1).Range("J" & i_curr + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("J").Key
            'wb_pr.Worksheets(1).Range("K" & j + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("K").Key
            wb_pr.Worksheets(1).Range("L" & i_curr + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("L").Key
            wb_pr.Worksheets(1).Range("M" & i_curr + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("M").Key
            wb_pr.Worksheets(1).Range("N" & i_curr + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("N").Key
            wb_pr.Worksheets(1).Range("O" & i_curr + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("O").Key
            
            If Not (mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("Comment") Is Nothing) Then
            add_comm mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("Comment").Key, wb_pr.Worksheets(1), i_curr + 20, 15, False
            wb_pr.Worksheets(1).Rows(20 + i_curr & ":" & 20 + i_curr).Interior.Color = 255
            End If
            
            '���ܱ������¼PR��
            'ws_partsingle.Cells(CInt(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Key), CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("PA_Index").Item(1).Key)) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("C").Key
            ws_partsingle.Cells(CLng(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CLng(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Key), CLng(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("PA_Index").Item(1).Key)) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("C").Key
            
            '���ܱ������¼PR��
End If


Next
'20180911 ���ô�ӡ����
wb_pr.Worksheets("PA").PageSetup.PrintArea = "$B$1:$O$" & i_curr + 20
'20180911 ���ô�ӡ����

wb_pr.Save
wb_pr.Close


Next

'�ڴ���PR������������PR�ļ�
''==============================


wb_fl.Save


'���ļ���
Shell "explorer.exe " & mokc.Item("FL").Item("FDN").Item("FDNPR").Key, vbNormalFocus
'���ļ���

Workbooks("PR_Create_Tool.xlsm").Saved = True
Workbooks("PR_Create_Tool.xlsm").Close


End Sub

'===========================
'li YiFei, 20170908 ���CostUnit �������:CN.505778.001,ϵͳ���Զ���41
Function COST_UNIT_check(str1 As String, Optional str2 As String = "") As String
If Len(str1) = 13 Then
If str1 Like "CN.??????.???" Then
str1 = Left(str1, 10) & "41" & Right(str1, 3)
End If
ElseIf str1 Like "?.?????.???" Then
If InStr(str2, "CN.") > 0 Then
str2 = Mid(str2, InStr(str2, "CN."), 9)
str1 = str2 & ".41" & Right(str1, 3)

End If

End If
COST_UNIT_check = str1
End Function
'li YiFei, 20170908 ���CostUnit �������:CN.505778.001,ϵͳ���Զ���41
'===========================



Sub PP_Creater()
Attribute PP_Creater.VB_ProcData.VB_Invoke_Func = "p\n14"
'�������ڴ������� PR����
'MsgBox "WinShuttle Can Use again! Please UPload to SAP after Create PR!"
back_followinglist
 If Enagble_addins("PPPP") Then
 End If
 
Dim wb_pr As Workbook

Dim mfso As New CFSO
Dim mokc As New OneKeyCls
mokc.Add "FL", "FL"
mokc.Add "PR", "PR"
mokc.Item("PR").Add "FLFP_TEMPLATE", "FLFP_TEMPLATE"
mokc.Item("PR").Add "OEM_NAME", "OEM_NAME"
mokc.Add "WS_PartSingle", "WS_PartSingle"
mokc.Item("WS_PartSingle").Add "WS_HEAD", "WS_HEAD"
mokc.Item("WS_PartSingle").Add "WS_BODY", "WS_BODY"


mokc.Item("FL").Add "FLN", "FLN"
mokc.Item("FL").Add "FDN", "FDN"

mokc.Item("PR").Add "PRN_LAST", "PRN_LAST"

mokc.Item("FL").Item("FDN").Add "FDNPR", "FDNPR"

mokc.Item("FL").Add "CUR_PR_NUM", "CUR_PR_NUM"


'=======================
'�ҵ��Ѿ��򿪵Ļ�е���ٱ�
Dim b_c As Boolean
Dim str1 As String
Dim str2 As String
Dim str3 As String
Dim str4 As String

Dim s_t1 As String
Dim s_t2 As String
Dim s_t3 As String
Dim s_t4 As String


Dim wb_fl As Workbook
Dim i As Long
Dim i_last As Long
Dim i_PRN_LAST As Integer
Dim s_PRN_LAST As String
Dim fln As String
Dim fdn As String
Dim ws_partsingle As Worksheet

Dim j As Long
Dim j_last As Long

'��ȡPartSingle����ĵ�Ԫ��
Dim POS As String
Dim QTY As String
Dim UNIT As String
Dim ItemName As String
Dim OEM_ID As String
Dim OEM_NAME As String
Dim TKID_SUBASS As String
Dim TKID_STATION As String
Dim PA_Index As String
Dim R_DATE As String
Dim E_DATE As String
Dim MATERIAL As String
Dim STANDARD As String
Dim DIMENSION As String
Dim dbl_qty As Double



'��ȡPartSingle����ĵ�Ԫ��




b_c = False
For i = 1 To Workbooks.Count
str1 = Workbooks(i).Name
If str1 Like "CN.*Pneumatics*Following*" Or str1 Like "CN.*Following*Pneumatics*" Then
Set wb_fl = Workbooks(i)
b_c = True
Exit For
End If
Next
If Not (wb_fl Is Nothing) Then
Msgbox "����PR�ĸ��ٱ�Ϊ��" & Chr(10) & wb_fl.Name
If wb_fl.ReadOnly = True Then
Msgbox "ֻ����ʽ�ĸ��ٱ��޷�����PR"
wb_fl.Close
Exit Sub
End If
Else
Msgbox "���ȴ򿪸��ٱ� CN.*Pneumatics*Following.xlsm "
Exit Sub
End If
'�ҵ��Ѿ��򿪵Ļ�е���ٱ�
'=======================
mokc.Item("FL").Item("FLN").Key = wb_fl.Name
mokc.Item("FL").Item("FDN").Key = wb_fl.Path
If Right(mokc.Item("FL").Item("FDN").Key, 1) <> "\" Then
mokc.Item("FL").Item("FDN").Key = mokc.Item("FL").Item("FDN").Key & "\"
End If
'=======================
'Checking  �ļ����Ƿ���� (mokc.Item("FL").Item("FDN").Item("FDNPR").Key)
fdn = mokc.Item("FL").Item("FDN").Key & "PAX_Mechanical&H&P"
mokc.Item("FL").Item("FDN").Item("FDNPR").Key = fdn
If mfso.folderexists(fdn) = False Then
If Msgbox("�ļ��в�����:" & Chr(10) & fdn & Chr(10) & "��Ҫ������ OK ", vbOKCancel) = vbOK Then
mfso.CreateFolder fdn
Msgbox "�Ѿ����� ���ڴ��PR���ļ���:" & Chr(10) & mokc.Item("FL").Item("FDN").Key & "PAX_Mechanical&H&P" & Chr(10) & "���ڸ��ļ���������һ��ȷ��PR��ʼ��ŵ�PR����Ĭ��0001��PR����"
Else
Exit Sub
End If
End If
'Checking  �ļ����Ƿ���� (mokc.Item("FL").Item("FDN").Item("FDNPR").Key)
'=======================







'=======================
'Get_Last PAE Number

i_PRN_LAST = 1
s_PRN_LAST = "0001"
Record_file_in_folder mokc.Item("FL").Item("FDN").Item("FDNPR"), mokc.Item("FL").Item("FDN").Item("FDNPR").Key, ".xlsm"
For i = 1 To mokc.Item("FL").Item("FDN").Item("FDNPR").Item("FILE").Count
fln = mokc.Item("FL").Item("FDN").Item("FDNPR").Item("FILE").Item(i).Item("FLN").Key
fdn = mokc.Item("FL").Item("FDN").Item("FDNPR").Item("FILE").Item(i).Item("FDN").Key
If fdn = mokc.Item("FL").Item("FDN").Item("FDNPR").Key Then
    If fln Like "P?####*.xlsm" Then
    str1 = Mid(fln, 3, 4)
        If CInt(str1) > CInt(s_PRN_LAST) Then
        s_PRN_LAST = str1
        i_PRN_LAST = CInt(s_PRN_LAST)
       
        
        
        End If
    ElseIf fln Like "PAE###*.xlsm" Then
    
    
        str1 = Mid(fln, 4, 3)
        If CInt(str1) > CInt(s_PRN_LAST) Then
        s_PRN_LAST = str1
        i_PRN_LAST = CInt(s_PRN_LAST)
       
        
        End If
    Else
    
    'MsgBox fln
    Application.StatusBar = fln

    End If
End If
Next
i_PRN_LAST = i_PRN_LAST + 1
s_PRN_LAST = CStr(i_PRN_LAST)
If Len(s_PRN_LAST) < 4 Then
s_PRN_LAST = Left("000", 4 - Len(s_PRN_LAST)) & s_PRN_LAST
End If
mokc.Item("PR").Item("PRN_LAST").Key = s_PRN_LAST
'Get_Last PAE Number
'=======================




'=====================
'��ȡ Part_single ���
b_c = True ' b_c = True  Part_Single ��ȡ�ɹ���b_c = False  Part_Single ��ȡʧ��
mokc.Item("WS_PartSingle").Add "M_C_P", "M_C_P"
'Record the pype: Controls or Mechanics or Pneumatics
'������ Parts_Single,ʧ��
If b_c Then
If ws_exist(wb_fl, "Parts_Single") = False Then
b_c = False
Msgbox "Following list ,���治���ڹ���� Parts_Single"
Else
Set ws_partsingle = wb_fl.Worksheets("Parts_Single")
End If
End If
'�޷��ж��� �� or �� or �� ʧ��
If b_c Then
If InStr(wb_fl.Name, "Mechanics") > 0 Then
mokc.Item("WS_PartSingle").Item("M_C_P").Key = "Mechanics"
ElseIf InStr(wb_fl.Name, "Controls") > 0 Then
mokc.Item("WS_PartSingle").Item("M_C_P").Key = "Controls"
ElseIf InStr(wb_fl.Name, "Pneumatics") > 0 Then
mokc.Item("WS_PartSingle").Item("M_C_P").Key = "Pneumatics"
Else
b_c = False
Msgbox wb_fl.Name & Chr(10) & "�����б���������µ���֮һ��Mechanics ��  Controls �� Pneumatics"
End If
End If
'�ж� Part_Single ��ʽ�Ƿ���Ԥ�ø�ʽ
If b_c Then
If mokc.Item("WS_PartSingle").Item("M_C_P").Key = "Controls" Then

'Template 1
'   If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 1, "Pos.", "POS") Then b_c = False
'    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 3, "Qty", "QTY") Then b_c = False
'    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 4, "Base Unit", "UNIT") Then b_c = False
'    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 6, "Matl. Descrip.", "ItemName") Then b_c = False
'    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 7, "Material No.", "TKID") Then b_c = False
'    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 8, "SPI", "SPI") Then b_c = False
'    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 11, "Manuf.Part.No.", "OEM_ID") Then b_c = False
'    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 12, "Basic Material", "MATERIAL") Then b_c = False
'    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 13, "Matl. Standard", "STANDARD") Then b_c = False
'    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 14, "Size/Dimension", "DIMENSION") Then b_c = False
'    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 15, "Manufacturer", "OEM_NAME") Then b_c = False
'    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 16, "Sub-Assy._Number", "TKID_SUBASS") Then b_c = False
'    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 17, "Station_Number", "TKID_STATION") Then b_c = False
'    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 18, "PA_Index", "PA_Index") Then If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 18, "PA_Number", "PA_Index") Then b_c = False
'    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 19, "Release_date", "R_DATE") Then b_c = False
'    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 20, "Expect Week", "E_DATE") Then b_c = False
'    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 21, "description", "DESC") Then If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 21, "Description", "DESC") Then b_c = False
'    If b_c = False Then MsgBox "Part_Single ��ͷ�޷�ʶ��"
    
    
    
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 7, 2, "Pos #", "POS") Then b_c = False
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 7, 4, "Qty", "QTY") Then b_c = False
    
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 7, 7, "Q" & Chr(10) & "C", "UNIT") Then b_c = False
    
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 7, 5, "Description", "ItemName") Then b_c = False
    
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 7, 3, "TK Ident Number", "TKID") Then b_c = False
    
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 7, 8, "SW", "SPI") Then b_c = False
    
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 7, 10, "Manufacturer Part Number", "OEM_ID") Then b_c = False
    
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 7, 9, "Material Name", "MATERIAL") Then b_c = False
    
    'If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 13, "Matl. Standard", "STANDARD") Then b_c = False
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 7, 6, " TECHNICAL PARAMETERS", "DIMENSION") Then b_c = False
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 7, 11, "Manufacturer", "OEM_NAME") Then b_c = False
    
    'If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 16, "Sub-Assy._Number", "TKID_SUBASS") Then b_c = False
    
    'If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 17, "Station_Number", "TKID_STATION") Then b_c = False
    
    '�����ɱ����ĺ���
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 7, 15, "Cost Unit", "WBS") Then b_c = False
    '�����ɱ����ĺ���
    
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 7, 16, "PA_Index", "PA_Index") Then b_c = False
    
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 7, 17, "Release_date", "R_DATE") Then b_c = False
    
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 7, 18, "Expect " & Chr(10) & "Week", "E_DATE") Then b_c = False
    
    
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 7, 20, "Description", "DESC") Then b_c = False
    If b_c = False Then Msgbox "Part_Single ��ͷ�޷�ʶ��"
    
ElseIf mokc.Item("WS_PartSingle").Item("M_C_P").Key = "Pneumatics" Then


    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 7, 2, "Pos #", "POS") Then b_c = False
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 7, 4, "Qty", "QTY") Then b_c = False
    
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 7, 7, "Q" & Chr(10) & "C", "UNIT") Then b_c = False
    
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 7, 5, "Description", "ItemName") Then b_c = False
    
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 7, 3, "TK Ident Number", "TKID") Then b_c = False
    
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 7, 8, "SW", "SPI") Then b_c = False
    
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 7, 10, "Manufacturer Part Number", "OEM_ID") Then b_c = False
    
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 7, 9, "Material Name", "MATERIAL") Then b_c = False
    
    'If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 13, "Matl. Standard", "STANDARD") Then b_c = False
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 7, 6, " TECHNICAL PARAMETERS", "DIMENSION") Then b_c = False
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 7, 11, "Manufacturer", "OEM_NAME") Then b_c = False
    
    'If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 16, "Sub-Assy._Number", "TKID_SUBASS") Then b_c = False
    
    'If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 17, "Station_Number", "TKID_STATION") Then b_c = False
    
    '�����ɱ����ĺ���
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 7, 15, "Cost Unit", "WBS") Then b_c = False
    '�����ɱ����ĺ���
    
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 7, 16, "PA_Index", "PA_Index") Then b_c = False
    
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 7, 17, "Release_date", "R_DATE") Then b_c = False
    
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 7, 18, "Expect " & Chr(10) & "Week", "E_DATE") Then b_c = False
    
    
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 7, 20, "Description", "DESC") Then b_c = False
    

Else
b_c = False
End If
End If


If b_c = False Then
Msgbox "�޷�ʶ��ı�ͷ��"
Exit Sub
End If





If b_c Then

i_last = ws_partsingle.UsedRange.Rows(ws_partsingle.UsedRange.Rows.Count).row
For i = 8 To i_last
'��ȡ������Ϊ�㣬PA_Index Ϊ�յ���




PA_Index = Trim(ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("PA_Index").Item(1).Key)))
QTY = Trim(ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("QTY").Item(1).Key)))


If Len(PA_Index) = 0 And Len(QTY) > 0 Then
OEM_NAME = Trim(ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("OEM_NAME").Item(1).Key)))
'TKID_STATION = Trim(ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("TKID_STATION").Item(1).Key)))
dbl_qty = 0
Str_TO_Dbl QTY, dbl_qty
If Not (dbl_qty > 0) Then
Msgbox "���޸���������������������кţ�" & i
b_c = False
End If
If b_c Then
If Len(OEM_NAME) = 0 Then
Msgbox "��Ӧ�����Ʋ���Ϊ�գ��кţ�" & i
b_c = False
End If
If b_c Then
mokc.Item("WS_PartSingle").Item("WS_BODY").Add CStr(i), CStr(i)
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("POS").Item(1).Key)), "POS"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("QTY").Item(1).Key)), "QTY"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("UNIT").Item(1).Key)), "UNIT"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("ItemName").Item(1).Key)), "ItemName"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("TKID").Item(1).Key)), "TKID"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("SPI").Item(1).Key)), "SPI"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("OEM_ID").Item(1).Key)), "OEM_ID"


mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("MATERIAL").Item(1).Key)), "MATERIAL"
'mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("STANDARD").Item(1).Key)), "STANDARD"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("DIMENSION").Item(1).Key)), "DIMENSION"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("OEM_NAME").Item(1).Key)), "OEM_NAME"
'mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("TKID_SUBASS").Item(1).Key)), "TKID_SUBASS"
'mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("TKID_STATION").Item(1).Key)), "TKID_STATION"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("PA_Index").Item(1).Key)), "PA_Index"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("R_DATE").Item(1).Key)), "R_DATE"
'add WBS
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("WBS").Item(1).Key)), "WBS"
'add WBS
E_DATE = format_date_DDMMYYYY(ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("E_DATE").Item(1).Key)))
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add E_DATE, "E_DATE"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("DESC").Item(1).Key)), "DESC"

End If
End If
End If
Next
End If
'��ȡ Part_single ���
'=====================

If b_c = False Then
Msgbox "��ȡ Part_Single ʧ�ܣ��޷�����PR"
Exit Sub
End If





'=====================
'���PR ģ���Ƿ���ڣ������ڿ���һ�ݵ���ǰĿ¼
'Z:\24_Temp\PA_Logs\V1.1\TEMPLATE\010c1612_Purchase Requisition.xlsm
If mfso.FileExists(mokc.Item("FL").Item("FDN").Key & "\010c1612_Purchase Requisition.xlsm") Then
Else
 If mfso.FileExists("Z:\24_Temp\PA_Logs\V1.1\TEMPLATE\010c1612_Purchase Requisition.xlsm") Then
 mfso.copy_file "Z:\24_Temp\PA_Logs\V1.1\TEMPLATE\010c1612_Purchase Requisition.xlsm", mokc.Item("FL").Item("FDN").Key & "\010c1612_Purchase Requisition.xlsm"
 Else
 Msgbox "�� PRģ�壺Z:\24_Temp\PA_Logs\V1.1\TEMPLATE\010c1612_Purchase Requisition.xlsm"
 b_c = False
 End If
End If
mokc.Item("PR").Item("FLFP_TEMPLATE").Key = mokc.Item("FL").Item("FDN").Key & "\010c1612_Purchase Requisition.xlsm"
'���PR ģ���Ƿ���ڣ������ڿ���һ�ݵ���ǰĿ¼
'=====================
If b_c = False Then
Msgbox "PR ģ�岻���ڣ��޷�����PR"
Exit Sub
End If






'==============================
'���ģ������Ŀ����Ŀ���Ƿ���ڣ�������Ҫ������
If open_wb(wb_pr, mokc.Item("PR").Item("FLFP_TEMPLATE").Key) Then
str1 = wb_pr.Worksheets(1).Range("G7")
Do While Len(str1) = 0
str1 = InputBox("��������Ŀ��", "PR ģ����Ϣ��д", "CN.#######")
If str1 = "CN.#######" Then str1 = ""
Loop
wb_pr.Worksheets(1).Range("G7") = str1
str1 = wb_pr.Worksheets(1).Range("M7")
Do While Len(str1) = 0
str1 = InputBox("��������Ŀ����", "PR ģ����Ϣ��д")
Loop
wb_pr.Worksheets(1).Range("M7") = str1
wb_pr.Save
wb_pr.Saved = True
wb_pr.Close
Else
b_c = False
End If
'���ģ������Ŀ����Ŀ���Ƿ���ڣ�������Ҫ������
'==============================
If b_c = False Then
Msgbox "PR ģ�� ����Ŀ���ƣ�����Ŀ��δ��д"
Exit Sub
End If







'============================
'��λ ���
For i = 1 To mokc.Item("WS_PartSingle").Item("WS_BODY").Count
str1 = UNIT_check(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Item("UNIT").Key)
If Len(str1) = 0 Then
b_c = False
Exit For
Else
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Item("UNIT").Key = str1
End If
Next

'��λ ���
'============================
If b_c = False Then
Msgbox "��λ��� ʧ�� " & Chr(10) & "�к�:" & mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Key
Exit Sub
End If



'============================
'�������� ���
For i = 1 To mokc.Item("WS_PartSingle").Item("WS_BODY").Count
str1 = mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Item("E_DATE").Key
If Len(str1) = 0 Then
b_c = False
Exit For
Else
End If
Next

'�������� ���
'============================
If b_c = False Then
Msgbox "�������ڼ�� ʧ�� " & Chr(10) & "�к�:" & mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Key
Exit Sub
End If




'============================
'COST_UNIT ���
For i = 1 To mokc.Item("WS_PartSingle").Item("WS_BODY").Count
str1 = COST_UNIT_check(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Item("WBS").Key, mokc.Item("FL").Item("FLN").Key)
If Len(str1) = 0 Then
b_c = False
Exit For
Else
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Item("WBS").Key = str1
End If
Next

'COST_UNIT ���
'============================
If b_c = False Then
Msgbox "�ɱ����ĺŶ�ȡ ʧ�� " & Chr(10) & "�к�:" & mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Key
Exit Sub
End If






'============================
'��Ӧ������ ��鼰����
For i = 1 To mokc.Item("WS_PartSingle").Item("WS_BODY").Count
str1 = OEM_NAME_check(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Item("OEM_NAME").Key)
If Len(str1) = 0 Then
b_c = False
Exit For
End If
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Item("OEM_NAME").Key = str1


If mokc.Item("PR").Item("OEM_NAME").Item(str1) Is Nothing Then
mokc.Item("PR").Item("OEM_NAME").Add str1, str1
mokc.Item("PR").Item("OEM_NAME").Item(str1).Add CStr(i), CStr(i)


Else
mokc.Item("PR").Item("OEM_NAME").Item(str1).Add CStr(i), CStr(i)
End If
Next
'��Ӧ������ ��鼰����
'============================
If b_c = False Then
Msgbox "��ȡ ��Ӧ�̼�� ʧ�ܣ��޷�����PR" & Chr(10) & "�к�:" & mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Key
Exit Sub
End If


'============================
'��׼�� �� ���� PA ����ֿ���
If mokc.Item("PR").Item("OEM_NAME").Count > 1 Then
If Not (mokc.Item("PR").Item("OEM_NAME").Item("NA") Is Nothing) Then
Msgbox "�кţ�" & mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item("NA").Item(1).Key)).Key
b_c = False
End If
End If
'��׼�� �� ���� PA ����ֿ���
'============================
If b_c = False Then
Msgbox "��׼��(NA)������ͷǱ�׼���ֿ���"
Exit Sub
End If



'=====================
'N/A ֱ�� ��д N/A
'��Ӧ�����Ʋ��ܰ��������ַ�
If mokc.Item("PR").Item("OEM_NAME").Count > 1 Then
If Not (mokc.Item("PR").Item("OEM_NAME").Item("N/A") Is Nothing) Then
For i = 1 To mokc.Item("PR").Item("OEM_NAME").Item("N/A").Count
ws_partsingle.Cells(CInt(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item("N/A").Item(i).Key)).Key), CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("PA_Index").Item(1).Key)) = "N/A"
Next
End If
End If

If Not (mokc.Item("PR").Item("OEM_NAME").Item("N/A") Is Nothing) Then mokc.Item("PR").Item("OEM_NAME").Remove "N/A"

For i = 1 To mokc.Item("PR").Item("OEM_NAME").Count
str1 = mokc.Item("PR").Item("OEM_NAME").Item(i).Key
If InStr(str1, "\") > 0 Then b_c = False
If InStr(str1, "/") > 0 Then b_c = False
If InStr(str1, "*") > 0 Then b_c = False
If InStr(str1, ":") > 0 Then b_c = False
If InStr(str1, "?") > 0 Then b_c = False
If b_c = False Then Exit For
Next
'N/A ֱ�� ��д N/A
'��Ӧ�����Ʋ��ܰ��������ַ�
'=====================
'���PR ģ���Ƿ���ڣ������ڿ���һ�ݵ���ǰĿ¼
If b_c = False Then
Msgbox "��Ӧ�����ư��������ַ�  \ / : * ? �޷�����PR�����޸ģ�" & str1
Exit Sub
End If





'=================================
'���Ӽ�(TKSE) ������ͼ��
If mokc.Item("PR").Item("OEM_NAME").Count > 0 Then
If Not (mokc.Item("PR").Item("OEM_NAME").Item("TKSE") Is Nothing) Then
For i = 1 To mokc.Item("PR").Item("OEM_NAME").Item("TKSE").Count
str1 = mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item("TKSE").Item(i).Key)).Item("TKID").Key
If Len(str1) = 0 Then
Msgbox "���Ӽ�������ͼ��,���ٱ��кţ�" & mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item("TKSE").Item(i).Key)).Key
b_c = False
End If
Next
End If
End If
'���Ӽ�(TKSE) ������ͼ��
'=================================
If b_c = False Then
Msgbox "���Ӽ�(TKSE)�޵�ɭͼ��,�޷�����PR"
Exit Sub
End If



'=================================
'��NA��N/A�����ͺź�TKID����ͬʱΪ��
For i = 1 To mokc.Item("WS_PartSingle").Item("WS_BODY").Count
str1 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Item("TKID").Key)
str2 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Item("OEM_ID").Key)
str3 = mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Item("OEM_NAME").Key
If (str3 <> "NA") Or (str3 <> "N/A") Then
If Len(str1) = 0 And Len(str2) = 0 Then
b_c = False
Msgbox "�Ǳ�����ͺţ���ɭ�Ų���ͬʱΪ�ա�����У�" & mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Key
Exit For
End If
End If
Next
'��NA��N/A�����ͺź�TKID����ͬʱΪ��
'=================================
If b_c = False Then
Msgbox "�����Ų���Ϊ��,�޷�����PR"
Exit Sub
End If

'20181021 �ͺ�����
sort_pr mokc
'20181021 �ͺ�����



'==============================
'�ڴ�������ȫ��PR����
For i = 1 To mokc.Item("PR").Item("OEM_NAME").Count
'�������PR��



'B
'SAP Item No.
'mokc.Item("PR").Item("OEM_NAME").Item(i).Add "B", "B"

For j = 1 To mokc.Item("PR").Item("OEM_NAME").Item(i).Count
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add CStr(j), "B"

'C
'Item ="PE00010001"
str1 = "PX" & mokc.Item("PR").Item("PRN_LAST").Key
str2 = CStr(j)


'str2 = Left("000", 4 - Len(str2)) & str2
'20180530 �޸ĺ����ʽPX00010001ΪPX0001.001
str2 = Left(".00", 4 - Len(str2)) & str2



str1 = str1 & str2
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str1, "C"


'D
'ShortText, ���Ӽ� TKID���⹺�� OEM_ID
'1.���Ӽ�=��TKID
'2.�⹺����ͬʱ���ͺţ�����TKID,��������D���ͺţ�TKID���������ݺϲ���MEMO��
'3.�⹺���������ͺ�
'4.�⹺��������TKID
'5.�⹺����û�ͺ�Ҳû��TKID
str1 = mokc.Item("PR").Item("OEM_NAME").Item(i).Key
If str1 = "TKSE" Then
str2 = mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("TKID").Key
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str2, "D"
Else
str2 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("TKID").Key)
str3 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("OEM_ID").Key)
If Len(str3) > 0 Then
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str3, "D"
ElseIf Len(str2) > 0 Then
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str2, "D"
End If
End If



'E
'�����Ĵ���Ϊ��
'str1 = mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("TKID_SUBASS").Key
'mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str1, "E"

'F
'ֱ�ӽ�OEM_NAME ����
str1 = mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("OEM_NAME").Key
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str1, "F"

'G
'����
'1.���Ӽ�.TKID**����
'2.�⹺�� �� ��쿹�ͨ����д ��������

str1 = mokc.Item("PR").Item("OEM_NAME").Item(i).Key
If str1 = "TKSE" Then
str2 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("TKID").Key)
str3 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("ItemName").Key)
If Len(str3) = 0 Then
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str2, "G"
Else
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str2 & "**" & str3, "G"
End If
Else
str3 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("DIMENSION").Key)
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str3, "G"
End If


'H
'CostUnit
'ʹ�� ���ٱ��������4λ CN.3  & ��λ������Ŀ�� & 41 & ��λ���ڹ�λ��

str3 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("WBS").Key)
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str3, "H"



'I
'����
str1 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("QTY").Key)
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str1, "I"


'J
'��λ
str1 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("UNIT").Key)
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str1, "J"


'L,M
'COST ELEMENT
'Other manufacturing material (Non-Independent Function) 40250000
's_str2 = "Other manufacturing material (Non-Independent Function)": s_str3 = "40250000": mokc.Item("PCE").Add s_str2, s_str2: mokc.Item("PCE").Item(s_str2).Add s_str3, s_str3
'Electrical  Parts Purchase  40270000
's_str2 = "Electrical  Parts Purchase": s_str3 = "40270000": mokc.Item("PCE").Add s_str2, s_str2: mokc.Item("PCE").Item(s_str2).Add s_str3, s_str3
'Pneumatic & Hydraulic   40280000
's_str2 = "Pneumatic & Hydraulic": s_str3 = "40280000": mokc.Item("PCE").Add s_str2, s_str2: mokc.Item("PCE").Item(s_str2).Add s_str3, s_str3
'Machinery & tooling (Single Part)   43202000
's_str2 = "Machinery & tooling (Single Part)": s_str3 = "43202000": mokc.Item("PCE").Add s_str2, s_str2: mokc.Item("PCE").Item(s_str2).Add s_str3, s_str3
If mokc.Item("WS_PartSingle").Item("M_C_P").Key = "Mechanics" Then
str1 = mokc.Item("PR").Item("OEM_NAME").Item(i).Key
If str1 = "TKSE" Then
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add "Machinery & tooling (Single Part)", "L"
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add "43202000", "M"
Else
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add "Other manufacturing material (Non-Independent Function)", "L"
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add "40250000", "M"
End If
ElseIf mokc.Item("WS_PartSingle").Item("M_C_P").Key = "Controls" Then
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add "Electrical  Parts Purchase", "L"
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add "40270000", "M"
ElseIf mokc.Item("WS_PartSingle").Item("M_C_P").Key = "Pneumatics" Then
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add "Pneumatic & Hydraulic", "L"
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add "40280000", "M"
End If





'N������
'�������ڸ�ʽת��
str1 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("E_DATE").Key)
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str1, "N"





'O��ע
'E�кϲ�I�кϲ�T��


str1 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("ItemName").Key) '5
str2 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("MATERIAL").Key) '9
str3 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("DESC").Key)
If InStr(str1, str2) > 0 Then
str1 = str1
ElseIf InStr(str2, str1) > 0 Then
str1 = str2
Else
str1 = str1 & str2
End If
If Len(str3) > 0 Then str1 = str1 & str3
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str1, "O"


'��� ����������Ϊ�գ�Ҳ����NAME��Ϊ�գ��򽫱�ע��Ų�뼼��������
str1 = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("G").Key
If Len(str1) = 0 Then
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("G").Key = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("O").Key
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("O").Key = ""
End If
'��� ����������Ϊ�գ�Ҳ����NAME��Ϊ�գ��򽫱�ע��Ų�뼼��������


Next






'PR���ż�1
i_PRN_LAST = CInt(mokc.Item("PR").Item("PRN_LAST").Key)
i_PRN_LAST = i_PRN_LAST + 1
s_PRN_LAST = CStr(i_PRN_LAST)
If Len(s_PRN_LAST) < 4 Then
s_PRN_LAST = Left("000", 4 - Len(s_PRN_LAST)) & s_PRN_LAST
End If
mokc.Item("PR").Item("PRN_LAST").Key = s_PRN_LAST
'PR���ż�1

 

Next
'�ڴ�������ȫ��PR����
'==============================


'====================================
'��Ԫ�񳤶Ƚ�����35������Ķ� mokc.Item("PR").Item("OEM_NAME") �Ĵ���
'D�г��� ��..��������O��
'G�г��� ��##��������O��
'O�г��� ��^^��ǰ��ķֿ���������� ע��
For i = 1 To mokc.Item("PR").Item("OEM_NAME").Count
For j = 1 To mokc.Item("PR").Item("OEM_NAME").Item(i).Count
s_t1 = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("D").Key
s_t2 = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("G").Key
s_t3 = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("O").Key
s_t4 = ""
If Len(s_t1) > 35 Then
s_t4 = ".." & Right(s_t1, Len(s_t1) - 33)
s_t1 = Left(s_t1, 33) & ".."
End If
If Len(s_t2) > 35 Then
s_t4 = s_t4 & "##" & Right(s_t2, Len(s_t2) - 33)
s_t2 = Left(s_t2, 33) & "##"
End If
If Len(s_t4) > 0 And Len(s_t3) > 0 Then
s_t4 = s_t4 & "^^" & s_t3
Else
s_t4 = s_t4 & s_t3
End If
If Len(s_t4) <= 35 Then
s_t3 = s_t4
s_t4 = ""
Else
s_t3 = Left(s_t4, 35)
s_t4 = Right(s_t4, Len(s_t4) - 35)
End If
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("D").Key = s_t1
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("G").Key = s_t2
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("O").Key = s_t3
If Len(s_t4) > 0 Then
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add s_t4, "Comment"
End If
Next
Next
'��Ԫ�񳤶Ƚ�����35������Ķ� mokc.Item("PR").Item("OEM_NAME") �Ĵ���
'====================================








'==============================
'�ڴ���PR������������PR�ļ�
'��ģ��


For i = 1 To mokc.Item("PR").Item("OEM_NAME").Count
'PR���ļ�����PX1515_CN.305587-8-9_Spare parts_20170725.xlsm
'PR���ļ�����PX####_CN.######_OEM_NAME_YYYYMMDD.xlsm
str1 = Left(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(1).Item("C").Key, 6)


str2 = Left(wb_fl.Name, 9)
str3 = mokc.Item("PR").Item("OEM_NAME").Item(i).Key
fln = str1 & "_" & str2 & "_H&P_" & str3 & "_" & Format(Now(), "YYYYMMDD") & ".xlsm"


open_wb wb_pr, mokc.Item("PR").Item("FLFP_TEMPLATE").Key
wb_pr.SaveAs mokc.Item("FL").Item("FDN").Item("FDNPR").Key & "\" & fln


'=======================================
'������Ԫ��Name of component .TK Internal Ident. number
wb_pr.Worksheets(1).Range("D20") = "Vendor Part No."
wb_pr.Worksheets(1).Range("G20") = "Name of component .TK Internal Ident. number"
'������Ԫ��Name of component .TK Internal Ident. number
'=======================================


wb_pr.Worksheets(1).Range("O7") = str1



'=========================Applicant:
 str1 = Application.UserName
If Len(str1) > 12 Then str1 = Environ("username")
If Len(str1) > 12 Then str1 = Left(str1, 12)
wb_pr.Worksheets(1).Range("C3") = str1
'=========================Applicant:


'=========================Application Date:
str1 = Format(Now(), "MM/DD/YYYY")
wb_pr.Worksheets(1).Range("M3") = str1
'=========================Application Date:



'��ͬWBS��SHORTTEXT ��Ҫ�ϲ���������ȡ�Ĳ����ǣ������ļ���д�����һ���� ����Ƿ���Ժϲ������Ժϲ���ϲ�
'i_PRN_LAST s_PRN_LAST ʵʱ����PR����д���ݣ�����ԭʼֵ

'��Ϊ���ShortText����35�ᱻ�ضϣ�����ԭ����ͬ���ͺţ�������Ϊ��ͬ���ϲ�����������Ҫȫ����ͬ�źϲ�����

i_PRN_LAST = 1
s_PRN_LAST = "1"
Dim i_curr As Integer

Dim WBS_s As String
Dim SortText_s As String
Dim Memo_s As String




For j = 1 To mokc.Item("PR").Item("OEM_NAME").Item(i).Count


'===
WBS_s = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("H").Key
SortText_s = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("D").Key
Memo_s = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("O").Key


OEM_NAME = mokc.Item("PR").Item("OEM_NAME").Item(i).Key
If mokc.Item("MERGE") Is Nothing Then mokc.Add "MERGE", "MERGE"
If mokc.Item("MERGE").Item(OEM_NAME) Is Nothing Then mokc.Item("MERGE").Add OEM_NAME, OEM_NAME

If mokc.Item("MERGE").Item(OEM_NAME).Item(WBS_s & SortText_s & Memo_s) Is Nothing Then

b_c = True
i_curr = i_PRN_LAST
i_PRN_LAST = i_PRN_LAST + 1


mokc.Item("MERGE").Item(OEM_NAME).Add WBS_s & SortText_s & Memo_s, WBS_s & SortText_s & Memo_s
mokc.Item("MERGE").Item(OEM_NAME).Item(WBS_s & SortText_s & Memo_s).Add CStr(i_curr), "PRN"

QTY = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("I").Key
mokc.Item("MERGE").Item(OEM_NAME).Item(WBS_s & SortText_s & Memo_s).Add QTY, "QTY"


str1 = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("C").Key
str2 = CStr(i_curr)
str2 = Left(".000", 4 - Len(str2)) & str2
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("C").Key = Left(str1, 6) & str2





Else
b_c = False
i_curr = CInt(mokc.Item("MERGE").Item(OEM_NAME).Item(WBS_s & SortText_s & Memo_s).Item("PRN").Key)
QTY = CStr(CInt(mokc.Item("MERGE").Item(OEM_NAME).Item(WBS_s & SortText_s & Memo_s).Item("QTY").Key) + CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("I").Key))
mokc.Item("MERGE").Item(OEM_NAME).Item(WBS_s & SortText_s & Memo_s).Item("QTY").Key = QTY


str1 = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("C").Key
str2 = CStr(i_curr)
str2 = Left(".000", 4 - Len(str2)) & str2
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("C").Key = Left(str1, 6) & str2




b_c = True
End If



'===

If b_c Then
            'wb_pr.Worksheets(1).Range("B" & i_curr + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("B").Key
            wb_pr.Worksheets(1).Range("B" & i_curr + 20) = i_curr
            
            wb_pr.Worksheets(1).Range("C" & i_curr + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("C").Key
            
            wb_pr.Worksheets(1).Range("D" & i_curr + 20).NumberFormat = "@"
            wb_pr.Worksheets(1).Range("D" & i_curr + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("D").Key
            'wb_pr.Worksheets(1).Range("E" & j + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("E").Key
            wb_pr.Worksheets(1).Range("F" & i_curr + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("F").Key
            wb_pr.Worksheets(1).Range("G" & i_curr + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("G").Key
            wb_pr.Worksheets(1).Range("H" & i_curr + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("H").Key
            
            'wb_pr.Worksheets(1).Range("I" & j + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("I").Key
            
            wb_pr.Worksheets(1).Range("I" & i_curr + 20) = QTY
            
            
            wb_pr.Worksheets(1).Range("J" & i_curr + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("J").Key
            'wb_pr.Worksheets(1).Range("K" & j + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("K").Key
            wb_pr.Worksheets(1).Range("L" & i_curr + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("L").Key
            wb_pr.Worksheets(1).Range("M" & i_curr + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("M").Key
            wb_pr.Worksheets(1).Range("N" & i_curr + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("N").Key
            wb_pr.Worksheets(1).Range("O" & i_curr + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("O").Key
            
            If Not (mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("Comment") Is Nothing) Then
            add_comm mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("Comment").Key, wb_pr.Worksheets(1), i_curr + 20, 15, False
            wb_pr.Worksheets(1).Rows(20 + i_curr & ":" & 20 + i_curr).Interior.Color = 255
            End If
            
            '���ܱ������¼PR��
            'ws_partsingle.Cells(CInt(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Key), CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("PA_Index").Item(1).Key)) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("C").Key
            ws_partsingle.Cells(CInt(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Key), CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("PA_Index").Item(1).Key)) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("C").Key
            
            '���ܱ������¼PR��
End If


Next

wb_pr.Save
wb_pr.Close


Next

'�ڴ���PR������������PR�ļ�
''==============================


wb_fl.Save


'���ļ���
Shell "explorer.exe " & mokc.Item("FL").Item("FDN").Item("FDNPR").Key, vbNormalFocus
'���ļ���

Workbooks("PR_Create_Tool.xlsm").Saved = True
Workbooks("PR_Create_Tool.xlsm").Close


End Sub


Sub PRNA_Creater()
Attribute PRNA_Creater.VB_ProcData.VB_Invoke_Func = "a\n14"

'��������������׼����PR���ӣ��ݶ���ĸ


Dim wb_pr As Workbook

Dim mfso As New CFSO
Dim mokc As New OneKeyCls
mokc.Add "FL", "FL"
mokc.Add "PR", "PR"
mokc.Item("PR").Add "FLFP_TEMPLATE", "FLFP_TEMPLATE"
mokc.Item("PR").Add "OEM_NAME", "OEM_NAME"
mokc.Add "WS_PartSingle", "WS_PartSingle"
mokc.Item("WS_PartSingle").Add "WS_HEAD", "WS_HEAD"
mokc.Item("WS_PartSingle").Add "WS_BODY", "WS_BODY"


mokc.Item("FL").Add "FLN", "FLN"
mokc.Item("FL").Add "FDN", "FDN"

mokc.Item("PR").Add "PRN_LAST", "PRN_LAST"

mokc.Item("FL").Item("FDN").Add "FDNPR", "FDNPR"

mokc.Item("FL").Add "CUR_PR_NUM", "CUR_PR_NUM"



Dim temp_s1 As String, temp_s2 As String, temp_s3 As String, temp_s4 As String


'=======================
'�ҵ��Ѿ��򿪵Ļ�е���ٱ�
Dim b_c As Boolean
Dim str1 As String
Dim str2 As String
Dim str3 As String
Dim str4 As String

Dim s_t1 As String
Dim s_t2 As String
Dim s_t3 As String
Dim s_t4 As String


Dim wb_fl As Workbook
Dim i As Long
Dim i_last As Long
Dim i_PRN_LAST As Integer
Dim s_PRN_LAST As String
Dim fln As String
Dim fdn As String
Dim ws_partsingle As Worksheet

Dim j As Long
Dim j_last As Long

'��ȡPartSingle����ĵ�Ԫ��
Dim POS As String
Dim QTY As String
Dim UNIT As String
Dim ItemName As String
Dim OEM_ID As String
Dim OEM_NAME As String
Dim TKID_SUBASS As String
Dim TKID_STATION As String
Dim PA_Index As String
Dim R_DATE As String
Dim E_DATE As String
Dim MATERIAL As String
Dim STANDARD As String
Dim DIMENSION As String
Dim dbl_qty As Double





'��ȡNA��
Dim ws_na As Worksheet
temp_s1 = "PR_Create_Tool.xlsm"
temp_s2 = "NA"
If mokc.Item("BZ") Is Nothing Then mokc.Add "BZ", "BZ"
If mokc.Item("XH") Is Nothing Then mokc.Add "XH", "XH"
Set ws_na = get_ws(Workbooks(temp_s1), temp_s2)
ws_na.Range("A1") = "BZ"
ws_na.Range("B1") = "BZ_STD"
ws_na.Range("C1") = "XH"
ws_na.Range("D1") = "XH_STD"
mokc_read_ws mokc.Item("BZ"), ws_na, 1, 2
mokc_read_ws mokc.Item("XH"), ws_na, 3, 4

'��ȡNA��



'��ȡPartSingle����ĵ�Ԫ��






b_c = False
For i = 1 To Workbooks.Count
str1 = Workbooks(i).Name
If str1 Like "CN.*Mechanics*Following*" Or str1 Like "CN.*Following*Mechanics*" Then
Set wb_fl = Workbooks(i)
b_c = True
Exit For
End If
Next
If Not (wb_fl Is Nothing) Then
Msgbox "����PR�ĸ��ٱ�Ϊ��" & Chr(10) & wb_fl.Name
If wb_fl.ReadOnly = True Then
Msgbox "ֻ����ʽ�ĸ��ٱ��޷�����PR"
wb_fl.Close
Exit Sub
End If
Else
Msgbox "���ȴ򿪸��ٱ� CN.*Mechanics*Following.xlsm "
Exit Sub
End If
'�ҵ��Ѿ��򿪵Ļ�е���ٱ�
'=======================
mokc.Item("FL").Item("FLN").Key = wb_fl.Name
mokc.Item("FL").Item("FDN").Key = wb_fl.Path
If Right(mokc.Item("FL").Item("FDN").Key, 1) <> "\" Then
mokc.Item("FL").Item("FDN").Key = mokc.Item("FL").Item("FDN").Key & "\"
End If
'=======================
'Checking  �ļ����Ƿ���� (mokc.Item("FL").Item("FDN").Item("FDNPR").Key)
fdn = mokc.Item("FL").Item("FDN").Key & "PAX_Mechanical&H&P"
mokc.Item("FL").Item("FDN").Item("FDNPR").Key = fdn
If mfso.folderexists(fdn) = False Then
If Msgbox("�ļ��в�����:" & Chr(10) & fdn & Chr(10) & "��Ҫ������ OK ", vbOKCancel) = vbOK Then
mfso.CreateFolder fdn
Msgbox "�Ѿ����� ���ڴ��PR���ļ���:" & Chr(10) & mokc.Item("FL").Item("FDN").Key & "PAX_Mechanical&H&P" & Chr(10) & "���ڸ��ļ���������һ��ȷ��PR��ʼ��ŵ�PR����Ĭ��0001��PR����"
Else
Exit Sub
End If
End If
'Checking  �ļ����Ƿ���� (mokc.Item("FL").Item("FDN").Item("FDNPR").Key)
'=======================







'=======================
'Get_Last PAX Number

i_PRN_LAST = 1
s_PRN_LAST = "0001"
Record_file_in_folder mokc.Item("FL").Item("FDN").Item("FDNPR"), mokc.Item("FL").Item("FDN").Item("FDNPR").Key, ".xlsm"
For i = 1 To mokc.Item("FL").Item("FDN").Item("FDNPR").Item("FILE").Count
fln = mokc.Item("FL").Item("FDN").Item("FDNPR").Item("FILE").Item(i).Item("FLN").Key
fdn = mokc.Item("FL").Item("FDN").Item("FDNPR").Item("FILE").Item(i).Item("FDN").Key
If fdn = mokc.Item("FL").Item("FDN").Item("FDNPR").Key Then
    If fln Like "P?####*.xlsm" Then
    str1 = Mid(fln, 3, 4)
        If CInt(str1) > CInt(s_PRN_LAST) Then
        s_PRN_LAST = str1
        i_PRN_LAST = CInt(s_PRN_LAST)
       
        
        
        End If
    ElseIf fln Like "PAX###*.xlsm" Then
    
    
        str1 = Mid(fln, 4, 3)
        If CInt(str1) > CInt(s_PRN_LAST) Then
        s_PRN_LAST = str1
        i_PRN_LAST = CInt(s_PRN_LAST)
       
        
        End If
    Else
    
    If Not (fln Like "MO*") Then
    Msgbox fln
    End If
    
    
    End If
End If
Next
i_PRN_LAST = i_PRN_LAST + 1
s_PRN_LAST = CStr(i_PRN_LAST)
If Len(s_PRN_LAST) < 4 Then
s_PRN_LAST = Left("000", 4 - Len(s_PRN_LAST)) & s_PRN_LAST
End If
mokc.Item("PR").Item("PRN_LAST").Key = s_PRN_LAST
'Get_Last PAX Number
'=======================




'=====================
'��ȡ Part_single ���
b_c = True ' b_c = True  Part_Single ��ȡ�ɹ���b_c = False  Part_Single ��ȡʧ��
mokc.Item("WS_PartSingle").Add "M_C_P", "M_C_P"
'Record the pype: Controls or Mechanics or Pneumatics
'������ Parts_Single,ʧ��
If b_c Then
If ws_exist(wb_fl, "Parts_Single") = False Then
b_c = False
Msgbox "Following list ,���治���ڹ���� Parts_Single"
Else
Set ws_partsingle = wb_fl.Worksheets("Parts_Single")
End If
End If
'�޷��ж��� �� or �� or �� ʧ��
If b_c Then
If InStr(wb_fl.Name, "Mechanics") > 0 Then
mokc.Item("WS_PartSingle").Item("M_C_P").Key = "Mechanics"
ElseIf InStr(wb_fl.Name, "Controls") > 0 Then
mokc.Item("WS_PartSingle").Item("M_C_P").Key = "Controls"
ElseIf InStr(wb_fl.Name, "Pneumatics") > 0 Then
mokc.Item("WS_PartSingle").Item("M_C_P").Key = "Pneumatics"
Else
b_c = False
Msgbox wb_fl.Name & Chr(10) & "�����б���������µ���֮һ��Mechanics ��  Controls �� Pneumatics"
End If
End If
'�ж� Part_Single ��ʽ�Ƿ���Ԥ�ø�ʽ
If b_c Then
If mokc.Item("WS_PartSingle").Item("M_C_P").Key = "Mechanics" Then

'Template 1
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 1, "Pos.", "POS") Then b_c = False
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 3, "Qty", "QTY") Then b_c = False
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 4, "Base Unit", "UNIT") Then b_c = False
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 6, "Matl. Descrip.", "ItemName") Then b_c = False
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 7, "Material No.", "TKID") Then b_c = False
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 8, "SPI", "SPI") Then b_c = False
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 11, "Manuf.Part.No.", "OEM_ID") Then b_c = False
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 12, "Basic Material", "MATERIAL") Then b_c = False
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 13, "Matl. Standard", "STANDARD") Then b_c = False
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 14, "Size/Dimension", "DIMENSION") Then b_c = False
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 15, "Manufacturer", "OEM_NAME") Then b_c = False
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 16, "Sub-Assy._Number", "TKID_SUBASS") Then b_c = False
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 17, "Station_Number", "TKID_STATION") Then b_c = False
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 18, "PA_Index", "PA_Index") Then If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 18, "PA_Number", "PA_Index") Then b_c = False
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 19, "Release_date", "R_DATE") Then b_c = False
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 20, "Expect Week", "E_DATE") Then b_c = False
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 21, "description", "DESC") Then If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 21, "Description", "DESC") Then b_c = False
    If Not TableHead_REC(mokc.Item("WS_PartSingle").Item("WS_HEAD"), ws_partsingle, 8, 23, "MO ID", "MO ID") Then b_c = False
    If b_c = False Then Msgbox "Part_Single ��ͷ�޷�ʶ��"
    
    
End If
End If



If b_c Then

i_last = ws_partsingle.UsedRange.Rows(ws_partsingle.UsedRange.Rows.Count).row
For i = 8 To i_last
'��ȡ������Ϊ�㣬PA_Index Ϊ�յ���




PA_Index = Trim(ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("PA_Index").Item(1).Key)))
QTY = Trim(ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("QTY").Item(1).Key)))


If Len(PA_Index) = 0 And Len(QTY) > 0 Then
OEM_NAME = Trim(ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("OEM_NAME").Item(1).Key)))
TKID_STATION = Trim(ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("TKID_STATION").Item(1).Key)))
dbl_qty = 0
Str_TO_Dbl QTY, dbl_qty
If Not (dbl_qty > 0) Then
Msgbox "���޸���������������������кţ�" & i
b_c = False
End If
If b_c Then
If Len(OEM_NAME) = 0 Then
Msgbox "��Ӧ�����Ʋ���Ϊ�գ��кţ�" & i
b_c = False
End If
If b_c Then
mokc.Item("WS_PartSingle").Item("WS_BODY").Add CStr(i), CStr(i)
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("POS").Item(1).Key)), "POS"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("QTY").Item(1).Key)), "QTY"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("UNIT").Item(1).Key)), "UNIT"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("ItemName").Item(1).Key)), "ItemName"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("TKID").Item(1).Key)), "TKID"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("SPI").Item(1).Key)), "SPI"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("OEM_ID").Item(1).Key)), "OEM_ID"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("MATERIAL").Item(1).Key)), "MATERIAL"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("STANDARD").Item(1).Key)), "STANDARD"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("DIMENSION").Item(1).Key)), "DIMENSION"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("OEM_NAME").Item(1).Key)), "OEM_NAME"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("TKID_SUBASS").Item(1).Key)), "TKID_SUBASS"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("TKID_STATION").Item(1).Key)), "TKID_STATION"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("PA_Index").Item(1).Key)), "PA_Index"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("R_DATE").Item(1).Key)), "R_DATE"

mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("MO ID").Item(1).Key)), "MO ID"

E_DATE = format_date_DDMMYYYY(ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("E_DATE").Item(1).Key)))
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add E_DATE, "E_DATE"
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CStr(i)).Add ws_partsingle.Cells(i, CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("DESC").Item(1).Key)), "DESC"

End If
End If
End If
Next
End If
'��ȡ Part_single ���
'=====================

If b_c = False Then
Msgbox "��ȡ Part_Single ʧ�ܣ��޷�����PR"
Exit Sub
End If





'=====================
'���PR ģ���Ƿ���ڣ������ڿ���һ�ݵ���ǰĿ¼
'Z:\24_Temp\PA_Logs\V1.1\TEMPLATE\010c1612_Purchase Requisition.xlsm
If mfso.FileExists(mokc.Item("FL").Item("FDN").Key & "\010c1612_Purchase Requisition.xlsm") Then
Else
 If mfso.FileExists("Z:\24_Temp\PA_Logs\V1.1\TEMPLATE\010c1612_Purchase Requisition.xlsm") Then
 mfso.copy_file "Z:\24_Temp\PA_Logs\V1.1\TEMPLATE\010c1612_Purchase Requisition.xlsm", mokc.Item("FL").Item("FDN").Key & "\010c1612_Purchase Requisition.xlsm"
 Else
 Msgbox "�� PRģ�壺Z:\24_Temp\PA_Logs\V1.1\TEMPLATE\010c1612_Purchase Requisition.xlsm"
 b_c = False
 End If
End If
mokc.Item("PR").Item("FLFP_TEMPLATE").Key = mokc.Item("FL").Item("FDN").Key & "\010c1612_Purchase Requisition.xlsm"
'���PR ģ���Ƿ���ڣ������ڿ���һ�ݵ���ǰĿ¼
'=====================
If b_c = False Then
Msgbox "PR ģ�岻���ڣ��޷�����PR"
Exit Sub
End If



'==============================
'���ģ������Ŀ����Ŀ���Ƿ���ڣ�������Ҫ������
If open_wb(wb_pr, mokc.Item("PR").Item("FLFP_TEMPLATE").Key) Then
str1 = wb_pr.Worksheets(1).Range("G7")
Do While Len(str1) = 0
str1 = InputBox("��������Ŀ��", "PR ģ����Ϣ��д", "CN.#######")
If str1 = "CN.#######" Then str1 = ""
Loop
wb_pr.Worksheets(1).Range("G7") = str1
str1 = wb_pr.Worksheets(1).Range("M7")
Do While Len(str1) = 0
str1 = InputBox("��������Ŀ����", "PR ģ����Ϣ��д")
Loop
wb_pr.Worksheets(1).Range("M7") = str1
wb_pr.Save
wb_pr.Saved = True
wb_pr.Close
Else
b_c = False
End If
'���ģ������Ŀ����Ŀ���Ƿ���ڣ�������Ҫ������
'==============================
If b_c = False Then
Msgbox "PR ģ�� ����Ŀ���ƣ�����Ŀ��δ��д"
Exit Sub
End If




'============================
'MO ��飬�����MO ����MO
Dim i_moid As Integer
i_moid = 0
For i = 1 To mokc.Item("WS_PartSingle").Item("WS_BODY").Count
str1 = UNIT_check(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Item("MO ID").Key)
If Len(str1) > 0 Then
If my_cint(str1) = 0 And i_moid > 0 Then
Msgbox "MO ���Ӻ� PX ���ӱ���ֿ��£�"
b_c = False
Exit For
End If
If i_moid < my_cint(str1) Then
i_moid = my_cint(str1)
If i_moid > 999 Then
Msgbox "MO ��ű���С��999!"
b_c = False
Exit For
End If
End If
End If
Next

'MO ��飬�����MO ����MO
'============================
If b_c = False Then
Msgbox "MO Check Error�� " & Chr(10) & "�к�:" & mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Key
Exit Sub
End If


'===================
'�����MO�����޸�
If i_moid > 0 Then
mokc.Item("PR").Item("PRN_LAST").Key = CStr(i_moid * 10)
mokc.Item("PR").Item("PRN_LAST").Key = Right("000" & mokc.Item("PR").Item("PRN_LAST").Key, 4)
End If
'�����MO�����޸�
'===================




'============================
'��λ ���
For i = 1 To mokc.Item("WS_PartSingle").Item("WS_BODY").Count
str1 = UNIT_check(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Item("UNIT").Key)
If Len(str1) = 0 Then
b_c = False
Exit For
Else
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Item("UNIT").Key = str1
End If
Next

'��λ ���
'============================
If b_c = False Then
Msgbox "��λ��� ʧ�� " & Chr(10) & "�к�:" & mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Key
Exit Sub
End If




'============================
'��Ӧ������ ��鼰����
For i = 1 To mokc.Item("WS_PartSingle").Item("WS_BODY").Count
str1 = OEM_NAME_check(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Item("OEM_NAME").Key)
If Len(str1) = 0 Then
b_c = False
Exit For
End If
mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Item("OEM_NAME").Key = str1


If mokc.Item("PR").Item("OEM_NAME").Item(str1) Is Nothing Then
mokc.Item("PR").Item("OEM_NAME").Add str1, str1
mokc.Item("PR").Item("OEM_NAME").Item(str1).Add CStr(i), CStr(i)


Else
mokc.Item("PR").Item("OEM_NAME").Item(str1).Add CStr(i), CStr(i)
End If
Next
'��Ӧ������ ��鼰����
'============================
If b_c = False Then
Msgbox "��ȡ ��Ӧ�̼�� ʧ�ܣ��޷�����PR" & Chr(10) & "�к�:" & mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Key
Exit Sub
End If


'============================
'��׼�� �� ���� PA ����ֿ���
If mokc.Item("PR").Item("OEM_NAME").Count = 2 And Not (mokc.Item("PR").Item("OEM_NAME").Item("N/A") Is Nothing) And Not (mokc.Item("PR").Item("OEM_NAME").Item("NA") Is Nothing) Then

ElseIf mokc.Item("PR").Item("OEM_NAME").Count > 1 Then

If Not (mokc.Item("PR").Item("OEM_NAME").Item("NA") Is Nothing) Then
Msgbox "�кţ�" & mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item("NA").Item(1).Key)).Key
b_c = False
End If
End If
'��׼�� �� ���� PA ����ֿ���
'============================
If b_c = False Then
Msgbox "��׼��(NA)������ͷǱ�׼���ֿ���"
Exit Sub
End If



'=====================
'N/A ֱ�� ��д N/A
'��Ӧ�����Ʋ��ܰ��������ַ�
If mokc.Item("PR").Item("OEM_NAME").Count > 1 Then
If Not (mokc.Item("PR").Item("OEM_NAME").Item("N/A") Is Nothing) Then
For i = 1 To mokc.Item("PR").Item("OEM_NAME").Item("N/A").Count
ws_partsingle.Cells(CInt(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item("N/A").Item(i).Key)).Key), CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("PA_Index").Item(1).Key)) = "N/A"
Next
End If
End If

If Not (mokc.Item("PR").Item("OEM_NAME").Item("N/A") Is Nothing) Then mokc.Item("PR").Item("OEM_NAME").Remove "N/A"

For i = 1 To mokc.Item("PR").Item("OEM_NAME").Count
str1 = mokc.Item("PR").Item("OEM_NAME").Item(i).Key
If InStr(str1, "\") > 0 Then b_c = False
If InStr(str1, "/") > 0 Then b_c = False
If InStr(str1, "*") > 0 Then b_c = False
If InStr(str1, ":") > 0 Then b_c = False
If InStr(str1, "?") > 0 Then b_c = False
If b_c = False Then Exit For
Next
'N/A ֱ�� ��д N/A
'��Ӧ�����Ʋ��ܰ��������ַ�
'=====================
'���PR ģ���Ƿ���ڣ������ڿ���һ�ݵ���ǰĿ¼
If b_c = False Then
Msgbox "��Ӧ�����ư��������ַ�  \ / : * ? �޷�����PR�����޸ģ�" & str1
Exit Sub
End If





'=================================
'���Ӽ�(TKSE) ������ͼ��
If mokc.Item("PR").Item("OEM_NAME").Count > 0 Then
If Not (mokc.Item("PR").Item("OEM_NAME").Item("TKSE") Is Nothing) Then
For i = 1 To mokc.Item("PR").Item("OEM_NAME").Item("TKSE").Count
str1 = mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item("TKSE").Item(i).Key)).Item("TKID").Key
If Len(str1) = 0 Then
Msgbox "���Ӽ�������ͼ��,���ٱ��кţ�" & mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item("TKSE").Item(i).Key)).Key
b_c = False
End If
Next
End If
End If
'���Ӽ�(TKSE) ������ͼ��
'=================================
If b_c = False Then
Msgbox "���Ӽ�(TKSE)�޵�ɭͼ��,�޷�����PR"
Exit Sub
End If



'=================================
'��NA��N/A�����ͺź�TKID����ͬʱΪ��
For i = 1 To mokc.Item("WS_PartSingle").Item("WS_BODY").Count
str1 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Item("TKID").Key)
str2 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Item("OEM_ID").Key)
str3 = mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Item("OEM_NAME").Key
If (str3 <> "NA") And (str3 <> "N/A") Then
If Len(str1) = 0 And Len(str2) = 0 Then
b_c = False
Msgbox "�Ǳ�����ͺţ���ɭ�Ų���ͬʱΪ�ա�����У�" & mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Key
Exit For
End If
End If
Next
'��NA��N/A�����ͺź�TKID����ͬʱΪ��
'=================================
If b_c = False Then
Msgbox "�����Ų���Ϊ��,�޷�����PR"
Exit Sub
End If



'=================================
'NA �����⴦��
If mokc.Item("PR").Item("OEM_NAME").Count = 1 Then


    '=================================
    'NA �� ��׼�Ų���Ϊ�գ�SIZE����Ϊ��
    For i = 1 To mokc.Item("WS_PartSingle").Item("WS_BODY").Count
    str1 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Item("STANDARD").Key)
    str2 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Item("DIMENSION").Key)
    str3 = mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Item("OEM_NAME").Key
    If str3 = "NA" Then
    If Len(str1) = 0 Or Len(str2) = 0 Then
    b_c = False
    Msgbox "�Ǳ�� ��׼�Ų���Ϊ�գ���׼�� �����Ϊ�ա�����У�" & mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Key
    Exit For
    End If
    End If
    Next
    'NA �� ��׼�Ų���Ϊ�գ�SIZE����Ϊ��
    '=================================
    If b_c = False Then
    Msgbox "�Ǳ�� NA ����д���淶(��׼�ţ�����Ϊ��) "
    Exit Sub
    End If
    
    
    '=================================
    'NA �� ��׼�Ÿ�ʽ����SIZE��ʽ��,���������
    For i = 1 To mokc.Item("WS_PartSingle").Item("WS_BODY").Count
    str1 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Item("STANDARD").Key)
    str2 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Item("DIMENSION").Key)
    str3 = mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Item("OEM_NAME").Key
    If str3 = "NA" Then
    
    If mokc.Item("BZ").Item("NA").Item("BODY").Item(str1) Is Nothing Then
    temp_s1 = InputBox("Standard  first use! please Confirm!" & Chr(10) & str1, "STANDARD CONFIRM!", Replace(str1, " ", ""))
    Do While Len(temp_s1) = 0
    temp_s1 = InputBox("Standard  first use! please Confirm!" & Chr(10) & str1, "STANDARD CONFIRM!", Replace(str1, " ", ""))
    Loop
    i_last = ws_na.UsedRange.Rows(ws_na.UsedRange.Rows.Count).row
    ws_na.Range("A" & i_last + 1) = str1
    ws_na.Range("B" & i_last + 1) = temp_s1
    mokc_read_ws_A mokc.Item("BZ"), ws_na
    Else
    str1 = mokc.Item("BZ").Item("NA").Item("BODY").Item(str1).Item(1).Key
    mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Item("STANDARD").Key = str1
    End If
    If mokc.Item("XH").Item("NA").Item("BODY").Item(str2) Is Nothing Then
    temp_s2 = InputBox("Size/Dimension  first use! please Confirm!" & Chr(10) & str2, "Size/Dimension!", Replace(str2, " ", ""))
    Do While Len(temp_s2) = 0
    temp_s2 = InputBox("Size/Dimension  first use! please Confirm!" & Chr(10) & str2, "Size/Dimension!", Replace(str2, " ", ""))
    Loop
    i_last = ws_na.UsedRange.Rows(ws_na.UsedRange.Rows.Count).row
    ws_na.Range("C" & i_last + 1) = str2
    ws_na.Range("D" & i_last + 1) = temp_s2
    mokc_read_ws_A mokc.Item("XH"), ws_na
    Else
    str2 = mokc.Item("XH").Item("NA").Item("BODY").Item(str2).Item(1).Key
    mokc.Item("WS_PartSingle").Item("WS_BODY").Item(i).Item("DIMENSION").Key = str2
    End If
    End If
    Next
    'NA �� ��׼�Ÿ�ʽ����SIZE��ʽ��
    '=================================

    '=================================
    'NA ���ͺ�ȡ ��׼��__��� NA�� ��ע ��д
    
    'NA ���ͺ�ȡ ��׼��__��� NA�� ��ע ��д
    '=================================
    
    
    
    
    
    '
    
    
End If
'NA �����⴦��
'=================================







'==============================
'�ڴ�������ȫ��PR����
For i = 1 To mokc.Item("PR").Item("OEM_NAME").Count
'�������PR��



'B
'SAP Item No.
'mokc.Item("PR").Item("OEM_NAME").Item(i).Add "B", "B"

For j = 1 To mokc.Item("PR").Item("OEM_NAME").Item(i).Count
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add CStr(j), "B"

'C
'Item ="PX00010001"
If i_moid = 0 Then
str1 = "PX" & mokc.Item("PR").Item("PRN_LAST").Key
Else
str1 = "MO" & mokc.Item("PR").Item("PRN_LAST").Key
End If

str2 = CStr(j)
str2 = Left(".00", 4 - Len(str2)) & str2
str1 = str1 & str2
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str1, "C"


'D
'ShortText, ���Ӽ� TKID���⹺�� OEM_ID
'1.���Ӽ�=��TKID
'2.�⹺����ͬʱ���ͺţ�����TKID,��������D���ͺţ�TKID���������ݺϲ���MEMO��
'3.�⹺���������ͺ�
'4.�⹺��������TKID
'5.�⹺����û�ͺ�Ҳû��TKID
str1 = mokc.Item("PR").Item("OEM_NAME").Item(i).Key
If str1 = "TKSE" Then
    str2 = mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("TKID").Key
    mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str2, "D"
ElseIf str1 = "NA" Then
    str2 = mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("STANDARD").Key & "__" & mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("DIMENSION").Key
    mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str2, "D"
Else
    str2 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("TKID").Key)
    str3 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("OEM_ID").Key)
    If Len(str3) > 0 Then
    mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str3, "D"
    ElseIf Len(str2) > 0 Then
    mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str2, "D"
    End If
End If



'E
'ֱ�ӽ�TKID_SUBASS����
str1 = mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("TKID_SUBASS").Key
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str1, "E"

'F
'ֱ�ӽ�OEM_NAME ����
str1 = mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("OEM_NAME").Key
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str1, "F"

'G
'����
'1.���Ӽ�.TKID**����
'2.�⹺��.����
str1 = mokc.Item("PR").Item("OEM_NAME").Item(i).Key
If str1 = "TKSE" Then
str2 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("TKID").Key)
str3 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("ItemName").Key)
If Len(str3) = 0 Then
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str2, "G"
Else
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str2 & "**" & str3, "G"
End If
Else
str3 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("ItemName").Key)
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str3, "G"
End If


'H
'CostUnit
'ʹ�� ���ٱ��������4λ CN.3  & ��λ������Ŀ�� & 41 & ��λ���ڹ�λ��
str1 = Left(wb_fl.Name, 4)
str2 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("TKID_STATION").Key)
str2 = Left(str2, 11)
str3 = str1 & Mid(str2, 3, 5) & ".41" & Right(str2, 3)
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str3, "H"



'I
'����
str1 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("QTY").Key)
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str1, "I"


'J
'��λ
str1 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("UNIT").Key)
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str1, "J"


'L,M
'COST ELEMENT
'Other manufacturing material (Non-Independent Function) 40250000
's_str2 = "Other manufacturing material (Non-Independent Function)": s_str3 = "40250000": mokc.Item("PCE").Add s_str2, s_str2: mokc.Item("PCE").Item(s_str2).Add s_str3, s_str3
'Electrical  Parts Purchase  40270000
's_str2 = "Electrical  Parts Purchase": s_str3 = "40270000": mokc.Item("PCE").Add s_str2, s_str2: mokc.Item("PCE").Item(s_str2).Add s_str3, s_str3
'Pneumatic & Hydraulic   40280000
's_str2 = "Pneumatic & Hydraulic": s_str3 = "40280000": mokc.Item("PCE").Add s_str2, s_str2: mokc.Item("PCE").Item(s_str2).Add s_str3, s_str3
'Machinery & tooling (Single Part)   43202000
's_str2 = "Machinery & tooling (Single Part)": s_str3 = "43202000": mokc.Item("PCE").Add s_str2, s_str2: mokc.Item("PCE").Item(s_str2).Add s_str3, s_str3
If mokc.Item("WS_PartSingle").Item("M_C_P").Key = "Mechanics" Then
str1 = mokc.Item("PR").Item("OEM_NAME").Item(i).Key
If str1 = "TKSE" Then
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add "Machinery & tooling (Single Part)", "L"
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add "43202000", "M"
Else
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add "Other manufacturing material (Non-Independent Function)", "L"
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add "40250000", "M"
End If
ElseIf mokc.Item("WS_PartSingle").Item("M_C_P").Key = "Controls" Then
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add "Electrical  Parts Purchase", "L"
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add "40270000", "M"
ElseIf mokc.Item("WS_PartSingle").Item("M_C_P").Key = "Pneumatics" Then
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add "Pneumatic & Hydraulic", "L"
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add "40280000", "M"
End If





'N������
'�������ڸ�ʽת��
str1 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("E_DATE").Key)
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str1, "N"





'O��ע
'���������ע
'1.���Ӽ���TKSE�������**Description
'2.�⹺��, TKID**���**Description
'3.NA �� TKID **Manuf.Part.No. & Basic Material
'4.MO �� Description + MO ����

str1 = mokc.Item("PR").Item("OEM_NAME").Item(i).Key
If str1 = "TKSE" Then
    str2 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("DIMENSION").Key)
    str3 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("DESC").Key)
    If Len(str2) = 0 Or Len(str3) = 0 Then
    mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str2 & str3, "O"
    Else
    mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str2 & "**" & str3, "O"
    End If
ElseIf str1 = "NA" Then
    str2 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("TKID").Key)
    If Len(str2) > 0 Then
    str2 = str2 & "**" & Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("OEM_ID").Key)
    Else
    str2 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("OEM_ID").Key)
    End If
    If InStr(str2, Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("MATERIAL").Key)) = 0 Then
    str2 = str2 & "**" & Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("MATERIAL").Key)
    End If
    If Right(str2, 2) = "**" Then str2 = Left(str2, Len(str2) - 2)
    mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str2, "O"
    

Else
str2 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("TKID").Key)
str3 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("DIMENSION").Key)
If Len(str2) = 0 Or Len(str3) = 0 Then
str2 = str2 & str3
Else
str2 = str2 & "**" & str3
End If
str3 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("DESC").Key)
If Len(str2) = 0 Or Len(str3) = 0 Then
str2 = str2 & str3
Else
str2 = str2 & "**" & str3
End If
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add str2, "O"
End If


If i_moid > 0 Then
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("O").Key = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("O").Key & "_MO" & Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("MO ID").Key)
End If





Next


'PR���ż�1
i_PRN_LAST = CInt(mokc.Item("PR").Item("PRN_LAST").Key)
i_PRN_LAST = i_PRN_LAST + 1
s_PRN_LAST = CStr(i_PRN_LAST)
If Len(s_PRN_LAST) < 4 Then
s_PRN_LAST = Left("000", 4 - Len(s_PRN_LAST)) & s_PRN_LAST
End If
mokc.Item("PR").Item("PRN_LAST").Key = s_PRN_LAST
'PR���ż�1

 

Next
'�ڴ�������ȫ��PR����
'==============================

'20181021 PR���ͺ�����
'sort_pr mokc

'20181021 PR���ͺ�����


'====================================
'��Ԫ�񳤶Ƚ�����35������Ķ� mokc.Item("PR").Item("OEM_NAME") �Ĵ���
'D�г��� ��..��������O��
'G�г��� ��##��������O��
'O�г��� ��^^��ǰ��ķֿ���������� ע��
For i = 1 To mokc.Item("PR").Item("OEM_NAME").Count
For j = 1 To mokc.Item("PR").Item("OEM_NAME").Item(i).Count
s_t1 = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("D").Key
s_t2 = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("G").Key
s_t3 = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("O").Key
s_t4 = ""
If Len(s_t1) > 35 Then
s_t4 = ".." & Right(s_t1, Len(s_t1) - 33)
s_t1 = Left(s_t1, 33) & ".."
End If
If Len(s_t2) > 35 Then
s_t4 = s_t4 & "##" & Right(s_t2, Len(s_t2) - 33)
s_t2 = Left(s_t2, 33) & "##"
End If
If Len(s_t4) > 0 And Len(s_t3) > 0 Then
s_t4 = s_t4 & "^^" & s_t3
Else
s_t4 = s_t4 & s_t3
End If
If Len(s_t4) <= 35 Then
s_t3 = s_t4
s_t4 = ""
Else
s_t3 = Left(s_t4, 35)
s_t4 = Right(s_t4, Len(s_t4) - 35)
End If
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("D").Key = s_t1
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("G").Key = s_t2
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("O").Key = s_t3
If Len(s_t4) > 0 Then
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Add s_t4, "Comment"
End If
Next
Next
'��Ԫ�񳤶Ƚ�����35������Ķ� mokc.Item("PR").Item("OEM_NAME") �Ĵ���
'====================================








'==============================
'�ڴ���PR������������PR�ļ�
'��ģ��


For i = 1 To mokc.Item("PR").Item("OEM_NAME").Count
'PR���ļ�����PX1515_CN.305587-8-9_Spare parts_20170725.xlsm
'PR���ļ�����PX####_CN.######_OEM_NAME_YYYYMMDD.xlsm
str1 = Left(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(1).Item("C").Key, 6)


str2 = Left(wb_fl.Name, 9)
str3 = mokc.Item("PR").Item("OEM_NAME").Item(i).Key
fln = str1 & "_" & str2 & "_" & str3 & "_" & Format(Now(), "YYYYMMDD") & ".xlsm"


open_wb wb_pr, mokc.Item("PR").Item("FLFP_TEMPLATE").Key
wb_pr.SaveAs mokc.Item("FL").Item("FDN").Item("FDNPR").Key & "\" & fln

wb_pr.Worksheets(1).Range("O7") = str1



'=======================================
'������Ԫ��Name of component .TK Internal Ident. number
wb_pr.Worksheets(1).Range("D20") = "Vendor Part No."
wb_pr.Worksheets(1).Range("G20") = "Name of component .TK Internal Ident. number"
'������Ԫ��Name of component .TK Internal Ident. number
'=======================================




'=========================Applicant:
 str1 = Application.UserName
If Len(str1) > 12 Then str1 = Environ("username")
If Len(str1) > 12 Then str1 = Left(str1, 12)
wb_pr.Worksheets(1).Range("C3") = str1
'=========================Applicant:


'=========================Application Date:
str1 = Format(Now(), "MM/DD/YYYY")
wb_pr.Worksheets(1).Range("M3") = str1
'=========================Application Date:



'��ͬWBS��SHORTTEXT ��Ҫ�ϲ���������ȡ�Ĳ����ǣ������ļ���д�����һ���� ����Ƿ���Ժϲ������Ժϲ���ϲ�
'i_PRN_LAST s_PRN_LAST ʵʱ����PR����д���ݣ�����ԭʼֵ
i_PRN_LAST = 1
s_PRN_LAST = "1"
Dim i_curr As Integer

Dim WBS_s As String
Dim SortText_s As String
Dim Memo_s As String




For j = 1 To mokc.Item("PR").Item("OEM_NAME").Item(i).Count


'===
'WBS_s = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("H").Key
WBS_s = ""
SortText_s = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("D").Key
Memo_s = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("O").Key



OEM_NAME = mokc.Item("PR").Item("OEM_NAME").Item(i).Key
If mokc.Item("MERGE") Is Nothing Then mokc.Add "MERGE", "MERGE"
If mokc.Item("MERGE").Item(OEM_NAME) Is Nothing Then mokc.Item("MERGE").Add OEM_NAME, OEM_NAME

If mokc.Item("MERGE").Item(OEM_NAME).Item(WBS_s & SortText_s & Memo_s) Is Nothing Then

b_c = True
i_curr = i_PRN_LAST
i_PRN_LAST = i_PRN_LAST + 1


mokc.Item("MERGE").Item(OEM_NAME).Add WBS_s & SortText_s & Memo_s, WBS_s & SortText_s & Memo_s
mokc.Item("MERGE").Item(OEM_NAME).Item(WBS_s & SortText_s & Memo_s).Add CStr(i_curr), "PRN"

QTY = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("I").Key
mokc.Item("MERGE").Item(OEM_NAME).Item(WBS_s & SortText_s & Memo_s).Add QTY, "QTY"


str1 = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("C").Key
str2 = CStr(i_curr)
str2 = Left(".000", 4 - Len(str2)) & str2
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("C").Key = Left(str1, 6) & str2





Else
b_c = False
i_curr = CInt(mokc.Item("MERGE").Item(OEM_NAME).Item(WBS_s & SortText_s & Memo_s).Item("PRN").Key)
QTY = CStr(CInt(mokc.Item("MERGE").Item(OEM_NAME).Item(WBS_s & SortText_s & Memo_s).Item("QTY").Key) + CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("I").Key))
mokc.Item("MERGE").Item(OEM_NAME).Item(WBS_s & SortText_s & Memo_s).Item("QTY").Key = QTY


str1 = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("C").Key
str2 = CStr(i_curr)
str2 = Left(".000", 4 - Len(str2)) & str2
mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("C").Key = Left(str1, 6) & str2




b_c = True
End If



'===

If b_c Then
            'wb_pr.Worksheets(1).Range("B" & i_curr + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("B").Key
            wb_pr.Worksheets(1).Range("B" & i_curr + 20) = i_curr
            wb_pr.Worksheets(1).Range("C" & i_curr + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("C").Key
            
            wb_pr.Worksheets(1).Range("D" & i_curr + 20).NumberFormat = "@"
            wb_pr.Worksheets(1).Range("D" & i_curr + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("D").Key
            'wb_pr.Worksheets(1).Range("E" & j + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("E").Key
            wb_pr.Worksheets(1).Range("F" & i_curr + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("F").Key
            wb_pr.Worksheets(1).Range("G" & i_curr + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("G").Key
            wb_pr.Worksheets(1).Range("H" & i_curr + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("H").Key
            
            'wb_pr.Worksheets(1).Range("I" & j + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("I").Key
            
            wb_pr.Worksheets(1).Range("I" & i_curr + 20) = QTY
            
            
            wb_pr.Worksheets(1).Range("J" & i_curr + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("J").Key
            'wb_pr.Worksheets(1).Range("K" & j + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("K").Key
            wb_pr.Worksheets(1).Range("L" & i_curr + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("L").Key
            wb_pr.Worksheets(1).Range("M" & i_curr + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("M").Key
            wb_pr.Worksheets(1).Range("N" & i_curr + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("N").Key
            wb_pr.Worksheets(1).Range("O" & i_curr + 20) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("O").Key
            
            If Not (mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("Comment") Is Nothing) Then
            add_comm mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("Comment").Key, wb_pr.Worksheets(1), i_curr + 20, 15, False
            wb_pr.Worksheets(1).Rows(20 + i_curr & ":" & 20 + i_curr).Interior.Color = 255
            End If
            
            '���ܱ������¼PR��
            'ws_partsingle.Cells(CInt(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Key), CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("PA_Index").Item(1).Key)) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("C").Key
            ws_partsingle.Cells(CInt(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Key), CInt(mokc.Item("WS_PartSingle").Item("WS_HEAD").Item("PA_Index").Item(1).Key)) = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("C").Key
            
                       '���ܱ������¼PR��
End If


Next

wb_pr.Save
wb_pr.Close


Next

'�ڴ���PR������������PR�ļ�
''==============================


wb_fl.Save


'���ļ���
Shell "explorer.exe " & mokc.Item("FL").Item("FDN").Item("FDNPR").Key, vbNormalFocus
'���ļ���


End Sub



Function sort_pr(mokc As OneKeyCls) As Boolean
'�������������PR�������ݵİ���������


 Dim i As Integer, j As Integer, k As Integer, i_min As Integer, kk As Integer
 
 Dim b_swi As Boolean
 Dim str1 As String, str2 As String, str3 As String
 Dim s_min As String
 
'Mechanics
'Controls
'Pneumatics
'mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Item("Comment")


 
 If mokc.Item("FL").Item("FLN").Key Like "*Mechanics*" Or mokc.Item("FL").Item("FLN").Key Like "*Controls*" Or mokc.Item("FL").Item("FLN").Key Like "*Pneumatics*" Then
 For i = 1 To mokc.Item("PR").Item("OEM_NAME").Count
 'ѭ��ȫ���Ĺ�Ӧ��
 For j = 1 To mokc.Item("PR").Item("OEM_NAME").Item(i).Count
 If mokc.Item("PR").Item("OEM_NAME").Item(i).Count > 2 Then
 '������3�����п�������
    'ѭ��ȫ������Ŀ
    i_min = j
    b_swi = False
    s_min = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("TKID").Key) & Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key)).Item("OEM_ID").Key)
    For k = j + 1 To mokc.Item("PR").Item("OEM_NAME").Item(i).Count
    '��ǰ���ͺ�����������Ƚϴ�С��������ֵ�ǰ������С�ģ��򲻶���������ֲ�����С�ģ������С�Ļ�����ǰλ��
    str1 = Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(k).Key)).Item("TKID").Key) & Trim(mokc.Item("WS_PartSingle").Item("WS_BODY").Item(CInt(mokc.Item("PR").Item("OEM_NAME").Item(i).Item(k).Key)).Item("OEM_ID").Key)
    If str1 < s_min Then
    s_min = str1
    str3 = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(k).Key
    i_min = k
    b_swi = True
    Else
    End If
    Next
    If b_swi = True Then
    mokc.Item("PR").Item("OEM_NAME").Item(i).Item(i_min).Key = mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key
    mokc.Item("PR").Item("OEM_NAME").Item(i).Item(j).Key = str3
    b_swi = False
    End If
End If
 
 Next
 Next
 
End If

End Function


Function Enagble_addins(str1 As String) As Boolean
Enagble_addins = True
On Error GoTo ERRORHAND
If Len(str1) = 0 Then
Exit Function
End If
Dim str2 As String
str2 = "Z:\24_Temp\PA_Logs\TOOLS\EXCEL_ADDIN\VISABLE\" & str1 & "#" & Application.UserName & ".txt"
Open str2 For Append As #1         ' ������ļ���
Print #1, Now()
Close #1
Const sAddinServerPath As String = "Z:\24_Temp\PA_Logs\TOOLS\EXCEL_ADDIN\MY_TOOL.xlam"
Dim sAddinLocalPath As String
Dim fs As Object
sAddinLocalPath = Application.UserLibraryPath & "MY_TOOL.xlam"
str2 = Dir(sAddinLocalPath)
If Len(str2) = 0 Then

If TypeName(get_addin("MY_TOOL.xlam")) <> "Nothing" Then
get_addin("MY_TOOL.xlam").Installed = False
End If



DoEvents
Set fs = CreateObject("Scripting.FileSystemObject")
fs.CopyFile sAddinServerPath, Application.UserLibraryPath, True
 Workbooks.Open sAddinLocalPath

End If

If TypeName(get_addin("MY_TOOL.xlam")) <> "Nothing" Then
get_addin("MY_TOOL.xlam").Installed = True
End If

Exit Function
ERRORHAND:
Enagble_addins = False
End Function

Function get_addin(fln As String) As AddIn
For Each get_addin In Application.AddIns
If fln = get_addin.Name Then
Exit For
End If

Next




End Function


Private Function Msgbox(str1 As String, Optional tf As Boolean = True) As VbMsgBoxResult
Application.StatusBar = str1
Msgbox = my_msgbox(str1 & Chr(10) & "See Status Bar!", tf)
End Function

Private Function Read_Main_to_PS(wb As Workbook) As Boolean

'��� �ж�ȡ��������ֹ ����PR���̣��ù���ʦ��� ��Main�����PartSingle������
Dim ws1 As Worksheet, ws2 As Worksheet
Dim i As Long, i_last As Long, j As Integer, j_last As Integer, k As Long, k_last As Long, ii As Long, jj As Long, kk As Long, ii_last As Long




Dim str1 As String, str2 As String, str3 As String, str4 As String, str5 As String, str6 As String, str7 As String, str8 As String, str9 As String, str10 As String


Dim date1 As Date, date2 As Date
Dim mokc As New OneKeyCls

Dim mtkbom As New CTKBOM


If ws_exist(wb, "BOM_Released_list") = False Then
Msgbox "ERROR: NO worksheet exist : BOM_Released_list "
Read_Main_to_PS = False
Exit Function
Else
Set ws1 = get_ws(wb, "BOM_Released_list")
End If
If ws_exist(wb, "Parts_Single") = False Then
Msgbox "ERROR: NO worksheet exist : Parts_Single"
Read_Main_to_PS = False
Exit Function
Else
Set ws2 = get_ws(wb, "Parts_Single")
End If

'��ʽ�����AB�� general ��H����
i_last = ws1.UsedRange.Rows(ws1.UsedRange.Rows.Count).row
ws1.Range("A9:B" & i_last).NumberFormat = "General"
ws1.Range("H9:H" & i_last).NumberFormat = "[$-409]d-mmm-yy;@"
ws1.Range("H9:H" & i_last).Interior.Color = RGB(255, 255, 255)


For i = 9 To i_last
str1 = ws1.Range("B" & i)
str2 = ws1.Range("G" & i)
str3 = format_date_DDMMYYYY(ws1.Range("H" & i))
If Len(str1) > 0 And Len(str2) = 0 Then
'��������� �޷�ʶ����ֹ����Ҫ����д������
If Len(str3) = 0 Then
Read_Main_to_PS = True
ws1.Range("H" & i).Interior.Color = RGB(255, 0, 0)
add_comment "Date Error", ws1.Range("H" & i)
'����������λ��
rg_scrowll ws1.Range("H" & i)
'����������λ��
Exit Function
ElseIf str3 Like "??.??.????" Then
date1 = CDate(Right(str3, 4) & "-" & Mid(str3, 4, 2) & "-" & Left(str3, 2))
date2 = Now()
If date1 - date2 < -1 Then
ws1.Range("H" & i).Interior.Color = RGB(255, 0, 0)
add_comment "Date Error:Date past!", ws1.Range("H" & i)
Read_Main_to_PS = True
'����������λ��
rg_scrowll ws1.Range("H" & i)
'����������λ��
Exit Function
End If
End If
mokc.Add CStr(i), CStr(i)
mokc.Item(CStr(i)).Add ws1.Range("B" & i), "B"
mokc.Item(CStr(i)).Add ws1.Range("C" & i), "C"
mokc.Item(CStr(i)).Add ws1.Range("D" & i), "D"
mokc.Item(CStr(i)).Add ws1.Range("F" & i), "F"
mokc.Item(CStr(i)).Add str3, "H"

End If
Next





'���BOM�Ƿ���ڣ����Ƿ����ʶ��
For i = 1 To mokc.Count
str1 = mokc.Item(i).Item("F").Key
If Right(str1, 1) <> "\" Then str1 = str1 & "\"
str2 = Dir(str1 & mokc.Item(i).Item("C").Key & "*.xls*")
If Len(str2) = 0 Then
Msgbox "FILE DOES NO EXIST! " & Chr(10) & str1 & mokc.Item(i).Item("C").Key & "*.xls*"
Read_Main_to_PS = True



Exit Function
Else
If mtkbom.read_bom_fl(str1 & str2) = True Then
'
'
Else
'
'
End If

'�д�����ֹ
If Len(mtkbom.Bom_Error(mokc.Item(i).Item("C").Key)) > 0 Then
'Msgbox "BOM CHECK ERROR:" & mokc.Item(i).Item("C").Key
ws1.Activate
ws1.Range("C" & mokc.Item(i).Key).Interior.Color = RGB(255, 0, 0)
add_comment mtkbom.Bom_Error(mokc.Item(i).Item("C").Key), ws1.Range("C" & mokc.Item(i).Key)

Read_Main_to_PS = True
Exit Function
End If
'�д�����ֹ


End If
Next
'���BOM�Ƿ���ڣ����Ƿ����ʶ��



'��д�У����ӣ�����

k_last = ws2.UsedRange.Rows(ws2.UsedRange.Rows.Count).row
ii_last = k_last + 1 '������ʼλ��


For i = 1 To mokc.Count

str1 = mokc.Item(i).Item("C").Key
'str1 ���׺ţ£ϣͣԣˣɣ�

j_last = mtkbom.get_bom_Body_Count(str1)
For j = 1 To j_last


str2 = mtkbom.get_bom_Body(str1, "bomitem_qty", j)
str3 = mtkbom.get_bom_Body(str1, "bomitem_manu", j)

If Len(str2) > 0 And Len(str3) > 0 Then

k_last = k_last + 1
'A~S �и�ֵ
str4 = mtkbom.get_bom_Body(str1, "bomitem_pos", j)
If Len(str4) > 0 Then
ws2.Range("A" & k_last) = str4
End If
str4 = mtkbom.get_bom_Body(str1, "bomitem_qty", j)
If Len(str4) > 0 Then
ws2.Range("C" & k_last) = CInt(str4) * CInt(mokc.Item(i).Item("B").Key)
End If
str4 = mtkbom.get_bom_Body(str1, "bomitem_unit", j)
If Len(str4) > 0 Then
ws2.Range("D" & k_last) = str4
End If
str4 = mtkbom.get_bom_Body(str1, "bomitem_desc", j)
If Len(str4) > 0 Then
ws2.Range("F" & k_last) = str4
End If
str4 = mtkbom.get_bom_Body(str1, "bomitem_tkid", j)
If Len(str4) > 0 Then
ws2.Range("G" & k_last) = str4
End If
str4 = mtkbom.get_bom_Body(str1, "bomitem_type", j)
If Len(str4) > 0 Then
ws2.Range("H" & k_last) = str4
End If
str4 = mtkbom.get_bom_Body(str1, "bomitem_manu_ref", j)
If Len(str4) > 0 Then
ws2.Range("K" & k_last) = str4
End If
str4 = mtkbom.get_bom_Body(str1, "bomitem_mat", j)
If Len(str4) > 0 Then
ws2.Range("M" & k_last) = str4
End If
str4 = mtkbom.get_bom_Body(str1, "bomitem_norm", j)
If Len(str4) > 0 Then
ws2.Range("M" & k_last) = str4
End If
str4 = mtkbom.get_bom_Body(str1, "bomitem_dime", j)
If Len(str4) > 0 Then
ws2.Range("N" & k_last) = str4
End If
str4 = mtkbom.get_bom_Body(str1, "bomitem_manu", j)
If Len(str4) > 0 Then
ws2.Range("O" & k_last) = str4
End If

str4 = mtkbom.get_bom_Body(str1, "bomitem_mcbf", j)
If Len(str4) > 0 Then
ws2.Range("I" & k_last) = str4
End If
str4 = mtkbom.get_bom_Body(str1, "bomitem_pc", j)
If Len(str4) > 0 Then
ws2.Range("J" & k_last) = str4
End If
str4 = mtkbom.get_bom_Body(str1, "bomitem_postype", j)
If Len(str4) > 0 Then
ws2.Range("B" & k_last) = str4
End If

str4 = mtkbom.get_bom_Body(str1, "bomitem_ver", j)
If Len(str4) > 0 Then
ws2.Range("E" & k_last) = str4
End If


ws2.Range("P" & k_last) = str1
str5 = mokc.Item(i).Item("D").Key
If str5 Like "?.?????.???*" Then
ws2.Range("Q" & k_last) = Left(str5, 11)
Else
ws2.Range("Q" & k_last) = str5
End If
ws2.Range("S" & k_last) = CDate(Format(Now(), "YYYY/mm/dd"))
str6 = mokc.Item(i).Item("H").Key
ws2.Range("T" & k_last) = CDate(Right(str6, 4) & "-" & Mid(str6, 4, 2) & "-" & Left(str6, 2))
'A~S �и�ֵ
Else
End If
Next

ws1.Range("G" & mokc.Item(i).Key) = "Done"

Next
'��д�У����ӣ�����


If k_last > ii_last Then ws2.Range("A" & ii_last & ":R" & k_last).NumberFormat = "General"

If ii_last = k_last + 1 Then
Read_Main_to_PS = False
Else
ws2.Activate
ws2.Range("A" & ii_last & ":T" & k_last).Select
Msgbox "Row " & ii_last & "TO " & k_last & "was add to partsingle List, Please Check!"
Read_Main_to_PS = True
End If



End Function

Private Function back_followinglist()
Dim mfso As New CFSO
Dim str1 As String, str2 As String, str3 As String
Dim wb As Workbook
If get_followinglist(wb) = False Then Exit Function
str1 = wb.FullName
str2 = wb.Name
str3 = "Z:\24_Temp\PA_Logs\PR\PR_Create_Tool\BACKUP\CW" & Application.WeekNum(Now()) & Format(Now(), "_YYYY") & "\"
If mfso.FileExists(str3 & str2) = False Then
'����
mfso.CreateFolder str3
mfso.copy_file str1, str3 & str2
'����
End If
End Function
Private Function get_followinglist(wb As Workbook) As Boolean
get_followinglist = False
For Each wb In Workbooks
If wb.Name Like "CN.*ollowing*.xlsm" Then
get_followinglist = True
Exit Function
End If
Next
End Function
Sub mac()
back_followinglist
End Sub
