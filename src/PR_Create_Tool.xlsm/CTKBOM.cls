VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTKBOM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mokc As New OneKeyCls

Private Sub Class_Initialize()
    
End Sub


Private Sub Class_Terminate()
Dim i As Integer
For i = 1 To mokc.Count
mokc.Remove (1)
Next
End Sub

Public Function read_bom_fl(FLFP As String) As Boolean
Dim str1 As String, str2 As String
str1 = Right(FLFP, (Len(FLFP) - InStrRev(FLFP, "\")))
str2 = Left(str1, 17)
Dim wb As Workbook
Dim TKID_BOM As String
Dim KEY_FLN_OR_HASH As String
TKID_BOM = str2
KEY_FLN_OR_HASH = str1
If mokc.Item(TKID_BOM) Is Nothing Then
'Set wb = Workbooks.Open(FLFP, 0, 1)
open_wb wb, FLFP
read_bom_wb wb
wb.Close 0
Else
If mokc.Item(TKID_BOM).Item(KEY_FLN_OR_HASH) Is Nothing Then
'Set wb = Workbooks.Open(FLFP, 0, 1)
open_wb wb, FLFP
read_bom_wb wb
wb.Close 0
End If
End If


End Function

Public Function get_bom_Head(TKID_BOM As String, bom_xxx As String, Optional KEY_FLN_OR_HASH As String = "") As String
'bom_type
'bom_tkid
'bom_desc
'bom_user1
On Error GoTo ERRORHAND
If mokc.Item(TKID_BOM) Is Nothing Then
get_bom_Head = ""
ElseIf mokc.Item(TKID_BOM).Item(KEY_FLN_OR_HASH) Is Nothing Then
get_bom_Head = mokc.Item(TKID_BOM).Item(1).Item("INF").Item(bom_xxx).Item("VAL").Key
Else
get_bom_Head = mokc.Item(TKID_BOM).Item(KEY_FLN_OR_HASH).Item("INF").Item(bom_xxx).Item("VAL").Key
End If
Exit Function
ERRORHAND:
get_bom_Head = ""
End Function

Public Function get_bom_Body(TKID_BOM As String, bomitem_xxx As String, row As Integer, Optional KEY_FLN_OR_HASH As String = "") As String
'bomitem_pos
'bomitem_start
'bomitem_tkid
'bomitem_qty
'bomitem_desc
'bomitem_manu
'bomitem_manu_ref


On Error GoTo ERRORHAND
If mokc.Item(TKID_BOM) Is Nothing Then
get_bom_Body = ""
ElseIf mokc.Item(TKID_BOM).Item(KEY_FLN_OR_HASH) Is Nothing Then
get_bom_Body = mokc.Item(TKID_BOM).Item(1).Item("BODY").Item(row).Item(bomitem_xxx).Key
Else
get_bom_Body = mokc.Item(TKID_BOM).Item(1).Item("BODY").Item(row).Item(bomitem_xxx).Key
End If
Exit Function
ERRORHAND:
get_bom_Body = ""
End Function



Public Function read_bom_wb(wb As Workbook, Optional TKID_BOM As String = "", Optional KEY_FLN_OR_HASH As String = "") As Boolean

'mokc.Item(D.xxxxx.xxx.ST.xx).Item(KEY_FLN_OR_HASH).Item ("INF")
'mokc.Item(D.xxxxx.xxx.ST.xx).Item(KEY_FLN_OR_HASH).Item ("BODY")
If Len(TKID_BOM) = 0 Then TKID_BOM = Left(wb.Name, 17)
If Len(KEY_FLN_OR_HASH) = 0 Then KEY_FLN_OR_HASH = wb.Name
If mokc.Item(TKID_BOM) Is Nothing Then
mokc.Add TKID_BOM, TKID_BOM
mokc.Item(TKID_BOM).Add KEY_FLN_OR_HASH, KEY_FLN_OR_HASH
Read_BOM wb, mokc.Item(TKID_BOM).Item(KEY_FLN_OR_HASH)
Else
If mokc.Item(TKID_BOM).Item(KEY_FLN_OR_HASH) Is Nothing Then
mokc.Item(TKID_BOM).Add KEY_FLN_OR_HASH, KEY_FLN_OR_HASH
Read_BOM wb, mokc.Item(TKID_BOM).Item(KEY_FLN_OR_HASH)
End If
End If
End Function

Private Function Read_BOM(wb As Workbook, mokc As OneKeyCls, Optional FLN_DATE_SIZE As String = "") As Boolean

'将BOM中内容读入 mokc
'清空mokc
Dim i As Integer, i_last As Integer
Dim j As Integer, j_last As Integer
Dim str1 As String, str2 As String, str3 As String, str4 As String, str5 As String, str6 As String


Dim ws As Worksheet

For i = 1 To mokc.Count
mokc.Remove (1)
Next
mokc.Add "INF", "INF"
mokc.Add "BODY", "BODY"
mokc.Item("INF").Add "FDN", "FDN"
str1 = wb.Path
mokc.Item("INF").Item("FDN").Add str1
str2 = wb.Name
mokc.Item("INF").Add "FLN", "FLN"
mokc.Item("INF").Item("FLN").Add str2

mokc.Item("INF").Add "FLN_DATE_SIZE", "FLN_DATE_SIZE"
mokc.Item("INF").Item("FLN_DATE_SIZE").Add FLN_DATE_SIZE





Set ws = wb.ActiveSheet



'检查ＢＯＭ是否可以识别

If ws.Range("A1") = "ThyssenKrupp Krause C.N" Then
    str1 = "TKSY_CN"
ElseIf (ws.Range("F4") = "           ThyssenKrupp System Engineering" Or ws.Range("F4") = "           ThyssenKrupp Krause GmbH") And ws.Range("O9") = "Manufacturer" Then
    str1 = "TKSY_SH"

'20180124 暂时取消

'ElseIf ws.Range("K1") = "Imprimé le :" Then
'    str1 = "TKSY_CN"
'20180124 暂时取消


ElseIf ws.Range("F2") = "STOCKLIST" Then
    str1 = "TKSY_ITA"
ElseIf ws.Range("C9") = "Qty" And ws.Range("F9") = "Matl. Descrip." And ws.Range("G9") = "Material No." And ws.Range("H9") = "SPI" And ws.Range("N9") = "Manufacturer" Then
    str1 = "TKSY_TES_V1"
ElseIf ws.Range("C9") = "Qty" And ws.Range("F9") = "Matl. Descrip." And ws.Range("G9") = "Material No." And ws.Range("H9") = "SPI" And ws.Range("O9") = "Manufacturer" Then
    str1 = "TKSY_TES_V2"
'20190508 庄工的特殊BOM
    If ws.Range("N2").MergeCells = True Then
    If ws.Range("N2").MergeArea.Address = "$N$2:$O$4" Then
    str1 = "TKSY_SH"
    End If
    End If
'20190508 庄工的特殊BOM
ElseIf ws.Range("A2") = "thyssenkrupp System Engineering" And ws.Range("F9") = "Material No." Then
    str1 = "TKSY_IND"
ElseIf ws.Range("A4") = "Pos" And ws.Range("D4") = "Material No" And ws.Range("M4") = "Manufacturer" Then
    str1 = "TKSY_PDM"
    
'增加TKSYSAS20180124
ElseIf ws.Range("F1") = "NOMENCLATURE" And ws.Range("B8") = "QTE" And ws.Range("D8") = "N° ARTICLE" And ws.Range("L8") = "FABRICANT" Then
    str1 = "TKSY_SAS"
ElseIf ws.Range("F1") = "                          PART LIST " And ws.Range("B8") = "QTE" And ws.Range("D8") = "N° " And ws.Range("L8") = "Manufacturer" Then
    str1 = "TKSY_SAS"

'增加TKSYSAS20180124
ElseIf ws.Range("F2") = "           ThyssenKrupp System Engineering" And ws.Range("A8") = "Pos." And ws.Range("C8") = "Qty" And ws.Range("F8") = "Matl. Descrip." And ws.Range("G8") = "Material No." And ws.Range("H8") = "SPI" And ws.Range("O8") = "Manufacturer" Then
'SURYA
str1 = "TKSY_SURYA"

'20190226
ElseIf ws.Range("A2") = "thyssenkrupp System Engineering" And ws.Range("F10") = "Material No." And ws.Range("S10") = "CNC Level" Then
str1 = "TKSY_SAP"
'20190226

Else
    str1 = "BOM_NOT_DEFINE"
End If
'检查ＢＯＭ是否可以识别



'给可以识别的ＢＯＭ赋地址

If str1 = "TKSY_SH" Then
'bom_type = "TKSY_SH": bom_tkid = "N2": bom_desc = "N5": bom_custid = ""
'bomitem_start = 10
'bom_ass_tkid = "G11"
'bomitem_pos = "A": bomitem_qty = "C": bomitem_tkid = "G": bomitem_type = "H": bomitem_desc = "F": bomitem_desc2 = "F": bomitem_manu = "O": bomitem_manu_ref = "K": bomitem_mat = "L": bomitem_norm = "M": bomitem_custid = "": bomitem_dime = "N"

str2 = "bom_type": str3 = "TKSY_SH": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bom_tkid": str3 = "N2": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
'Add User_C
str2 = "bom_user1": str3 = "D3": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
'Add User_C
str2 = "bom_desc": str3 = "N5": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
'str2 = "bom_custid": str3 = "": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_start": str3 = "10": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bom_ass_tkid": str3 = "G11": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_pos": str3 = "A": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_qty": str3 = "C": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_tkid": str3 = "G": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_type": str3 = "H": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_desc": str3 = "F": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_desc2": str3 = "F": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_manu": str3 = "O": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_manu_ref": str3 = "K": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_mat": str3 = "L": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_norm": str3 = "M": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
'str2 = "bomitem_custid": str3 = "": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_dime": str3 = "N": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"

'20190226
str2 = "bomitem_unit": str3 = "D": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"

str2 = "bomitem_mcbf": str3 = "I": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_pc": str3 = "J": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_postype": str3 = "B": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_ver": str3 = "E": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"

'20190226

ElseIf str1 = "TKSY_IND" Then
'

str2 = "bom_type": str3 = "TKSY_IND": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bom_tkid": str3 = "G5": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
'Add User_C
str2 = "bom_user1": str3 = "Q4": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
'Add User_C
str2 = "bom_desc": str3 = "G2": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
'str2 = "bom_custid": str3 = "": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_start": str3 = "10": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bom_ass_tkid": str3 = "F10": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_pos": str3 = "A": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_qty": str3 = "C": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_tkid": str3 = "F": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_type": str3 = "J": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_desc": str3 = "E": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
'str2 = "bomitem_desc2": str3 = "F": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_manu": str3 = "P": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_manu_ref": str3 = "Q": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_mat": str3 = "M": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_norm": str3 = "N": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
'str2 = "bomitem_custid": str3 = "": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_dime": str3 = "O": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"

ElseIf str1 = "TKSY_PDM" Then
'

str2 = "bom_type": str3 = "TKSY_PDM": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bom_tkid": str3 = "A1": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
'Add User_C
str2 = "bom_user1": str3 = "A3": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
'Add User_C
str2 = "bom_desc": str3 = "F5": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
'str2 = "bom_custid": str3 = "": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_start": str3 = "5": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bom_ass_tkid": str3 = "D5": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_pos": str3 = "A": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_qty": str3 = "C": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_tkid": str3 = "D": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_type": str3 = "J": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_desc": str3 = "F": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
'str2 = "bomitem_desc2": str3 = "F": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_manu": str3 = "M": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_manu_ref": str3 = "G": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
'str2 = "bomitem_mat": str3 = "M": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_norm": str3 = "H": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
'str2 = "bomitem_custid": str3 = "": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_dime": str3 = "I": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"




ElseIf str1 = "TKSY_SAP" Then

'bom_ass_tkid = "G6"
'bom_type = "TKSY_ITA": bom_tkid = "D6": bom_desc = "F6": bom_custid = "G7"
'bomitem_start = 9
'bomitem_pos = "A": bomitem_qty = "C": bomitem_tkid = "E": bomitem_type = "B": bomitem_desc = "F": bomitem_desc2 = "G": bomitem_manu = "J": bomitem_manu_ref = "H": bomitem_norm = "I": bomitem_custid = "D": bomitem_dime = ""
str2 = "bom_type": str3 = "TKSY_SAP": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bom_tkid": str3 = "G2": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bom_desc": str3 = "A4": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
'str2 = "bom_custid": str3 = "G7": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_start": str3 = "11": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bom_ass_tkid": str3 = "G2": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_pos": str3 = "A": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_qty": str3 = "C": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_tkid": str3 = "F": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_type": str3 = "J": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_desc": str3 = "E": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
'str2 = "bomitem_desc2": str3 = "G": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_manu": str3 = "P": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_manu_ref": str3 = "Q": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_mat": str3 = "M": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_norm": str3 = "N": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
'str2 = "bomitem_custid": str3 = "D": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_dime": str3 = "O": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"


'20190226
str2 = "bomitem_unit": str3 = "D": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_mcbf": str3 = "K": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
'str2 = "bomitem_pc": str3 = "J": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
'str2 = "bomitem_postype": str3 = "B": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_ver": str3 = "H": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"

'20190226



ElseIf str1 = "TKSY_ITA" Then
'bom_ass_tkid = "G6"
'bom_type = "TKSY_ITA": bom_tkid = "D6": bom_desc = "F6": bom_custid = "G7"
'bomitem_start = 9
'bomitem_pos = "A": bomitem_qty = "C": bomitem_tkid = "E": bomitem_type = "B": bomitem_desc = "F": bomitem_desc2 = "G": bomitem_manu = "J": bomitem_manu_ref = "H": bomitem_norm = "I": bomitem_custid = "D": bomitem_dime = ""
str2 = "bom_type": str3 = "TKSY_ITA": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bom_tkid": str3 = "D6": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bom_desc": str3 = "F6": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bom_custid": str3 = "G7": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_start": str3 = "9": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bom_ass_tkid": str3 = "G6": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_pos": str3 = "A": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_qty": str3 = "C": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_tkid": str3 = "E": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_type": str3 = "B": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_desc": str3 = "F": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_desc2": str3 = "G": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_manu": str3 = "J": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_manu_ref": str3 = "H": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
'str2 = "bomitem_mat": str3 = "L": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_norm": str3 = "I": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_custid": str3 = "D": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
'str2 = "bomitem_dime": str3 = "": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
ElseIf str1 = "TKSY_TES_V1" Then
'冷试设备上有V.06843.300.ST.00,第一次遇到
'ElseIf temp_str = "TKSY_TES_V1" Then
'bom_ass_tkid = "G11"
'bom_type = "TKSY_TES_V1": bom_tkid = "M2": bom_desc = "M5": bom_custid = ""
'bomitem_start = 11
'bomitem_pos = "A": bomitem_qty = "C": bomitem_tkid = "G": bomitem_type = "H": bomitem_desc = "F": bomitem_desc2 = "F": bomitem_manu = "N": bomitem_manu_ref = "J": bomitem_norm = "L": bomitem_custid = "": bomitem_dime = "M"
str2 = "bom_type": str3 = "TKSY_TES_V1": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bom_tkid": str3 = "M2": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bom_desc": str3 = "M5": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
'str2 = "bom_custid": str3 = "": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_start": str3 = "11": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bom_ass_tkid": str3 = "G11": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_pos": str3 = "A": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_qty": str3 = "C": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_tkid": str3 = "G": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_type": str3 = "H": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_desc": str3 = "F": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_desc2": str3 = "F": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_manu": str3 = "N": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_manu_ref": str3 = "J": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
'str2 = "bomitem_mat": str3 = "L": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_norm": str3 = "L": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
'str2 = "bomitem_custid": str3 = "": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_dime": str3 = "M": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"


ElseIf str1 = "TKSY_SAS" Then
'bom_type = "TKSY_SH": bom_tkid = "N2": bom_desc = "N5": bom_custid = ""
'bomitem_start = 10
'bom_ass_tkid = "G11"
'bomitem_pos = "A": bomitem_qty = "C": bomitem_tkid = "G": bomitem_type = "H": bomitem_desc = "F": bomitem_desc2 = "F": bomitem_manu = "O": bomitem_manu_ref = "K": bomitem_mat = "L": bomitem_norm = "M": bomitem_custid = "": bomitem_dime = "N"

str2 = "bom_type": str3 = "TKSY_SAS": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bom_tkid": str3 = "D4": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
'Add User_C
str2 = "bom_user1": str3 = "L2": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
'Add User_C
str2 = "bom_desc": str3 = "G2": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
'str2 = "bom_custid": str3 = "": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_start": str3 = "9": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bom_ass_tkid": str3 = "D9": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_pos": str3 = "A": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_qty": str3 = "B": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_tkid": str3 = "D": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_type": str3 = "E": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_desc": str3 = "F": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_desc2": str3 = "G": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_manu": str3 = "L": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_manu_ref": str3 = "H": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_mat": str3 = "J": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_norm": str3 = "K": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
'str2 = "bomitem_custid": str3 = "": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_dime": str3 = "I": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"





ElseIf str1 = "TKSY_TES_V2" Then
'冷试设备上有 在 TKSY_TES_V1 基础上 增加 J列 PC 新增一种 BOM格式
'ElseIf temp_str = "TKSY_TES_V2" Then
'bom_ass_tkid = "G11"
'bom_type = "TKSY_TES_V2": bom_tkid = "M2": bom_desc = "N5": bom_custid = ""
'bomitem_start = 11
'bomitem_pos = "A": bomitem_qty = "C": bomitem_tkid = "G": bomitem_type = "H": bomitem_desc = "F": bomitem_desc2 = "F": bomitem_manu = "O": bomitem_manu_ref = "K": bomitem_norm = "M": bomitem_custid = "": bomitem_dime = "N"
str2 = "bom_type": str3 = "TKSY_TES_V2": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3
str2 = "bom_tkid": str3 = "M2": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bom_desc": str3 = "N5": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
'str2 = "bom_custid": str3 = "": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_start": str3 = "11": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bom_ass_tkid": str3 = "G11": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_pos": str3 = "A": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_qty": str3 = "C": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_tkid": str3 = "G": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_type": str3 = "H": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_desc": str3 = "F": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_desc2": str3 = "F": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_manu": str3 = "O": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_manu_ref": str3 = "K": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_mat": str3 = "L": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_norm": str3 = "M": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
'str2 = "bomitem_custid": str3 = "": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_dime": str3 = "N": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"



ElseIf str1 = "TKSY_SURYA" Then
'CJLR P3,印度使用类似 TKSY_TES_V2格式
'ElseIf temp_str = "TKSY_TES_V2" Then
'bom_ass_tkid = "G11"
'bom_type = "TKSY_TES_V2": bom_tkid = "M2": bom_desc = "N5": bom_custid = ""
'bomitem_start = 11
'bomitem_pos = "A": bomitem_qty = "C": bomitem_tkid = "G": bomitem_type = "H": bomitem_desc = "F": bomitem_desc2 = "F": bomitem_manu = "O": bomitem_manu_ref = "K": bomitem_norm = "M": bomitem_custid = "": bomitem_dime = "N"
str2 = "bom_type": str3 = "TKSY_SURYA": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3
str2 = "bom_tkid": str3 = "N2": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bom_desc": str3 = "N5": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
'str2 = "bom_custid": str3 = "": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_start": str3 = "9": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bom_ass_tkid": str3 = "G9": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_pos": str3 = "A": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_qty": str3 = "C": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_tkid": str3 = "G": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_type": str3 = "H": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_desc": str3 = "F": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_desc2": str3 = "F": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_manu": str3 = "O": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_manu_ref": str3 = "K": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_mat": str3 = "L": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_norm": str3 = "M": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
'str2 = "bomitem_custid": str3 = "": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"
str2 = "bomitem_dime": str3 = "N": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"


Else

str2 = "bom_type": str3 = "BOM_NOT_DEFINE": mokc.Item("INF").Add str2, str2: mokc.Item("INF").Item(str2).Add str3, "ADDRESS"

Read_BOM = False
Exit Function
End If
'给可以识别的ＢＯＭ赋地址


'给可以识别的ＢＯＭ赋　值
If mokc.Item("INF").Item("bom_type").Item(1).Key <> "BOM_NOT_DEFINE" Then


str1 = "bom_tkid"
If Not (mokc.Item("INF").Item(str1) Is Nothing) Then
str2 = ws.Range(mokc.Item("INF").Item(str1).Item("ADDRESS").Key)
mokc.Item("INF").Item("bom_tkid").Add str2, "VAL"
End If


str1 = "bom_user1"
If Not (mokc.Item("INF").Item(str1) Is Nothing) Then
str2 = ws.Range(mokc.Item("INF").Item(str1).Item("ADDRESS").Key)
mokc.Item("INF").Item("bom_user1").Add str2, "VAL"
End If


str1 = "bom_desc"
If Not (mokc.Item("INF").Item(str1) Is Nothing) Then
str2 = ws.Range(mokc.Item("INF").Item(str1).Item("ADDRESS").Key)
mokc.Item("INF").Item("bom_desc").Add str2, "VAL"
End If

str1 = "bom_ass_tkid"
If Not (mokc.Item("INF").Item(str1) Is Nothing) Then
str2 = ws.Range(mokc.Item("INF").Item(str1).Item("ADDRESS").Key)
'如果BOM TKID like "?.?????.???.ST.??.??",或者不包含“ 00.00.00",则BOM
If (str2 Like "?.?????.???.ST.??") Or (Not (str2 Like "?.?????.???.00.??")) Then
str2 = Replace(mokc.Item("INF").Item("bom_tkid").Item("VAL").Key, "ST", "00")
End If
'如果BOM TKID like "?.?????.???.ST.??.??",或者不包含“ 00.00.00",则BOM
mokc.Item("INF").Item("bom_ass_tkid").Add str2, "VAL"
End If




'==检查 DIFF_BOMTKID 错误

If Get_FlnTkid(mokc.Item("INF").Item("FLN").Item(1).Key) <> mokc.Item("INF").Item("bom_tkid").Item("VAL").Key Then
If mokc.Item("ERROR") Is Nothing Then mokc.Add "ERROR", "ERROR"
If mokc.Item("ERROR").Item("DIFF_BOMTKID") Is Nothing Then mokc.Item("ERROR").Add "DIFF_BOMTKID", "DIFF_BOMTKID"
mokc.Item("ERROR").Item("DIFF_BOMTKID").Add mokc.Item("INF").Item("FLN").Item(1).Key & "_" & mokc.Item("INF").Item("bom_tkid").Item("VAL").Key
End If

'==检查 DIFF_BOMTKID 错误





i_last = ws.UsedRange.Rows(ws.UsedRange.Rows.Count).row
j = 0
For i = CInt(mokc.Item("INF").Item("bomitem_start").Item(1).Key) To i_last

'有删除线或者隐藏行，报错并跳过


'有删除线或者隐藏行，报错并跳过




str2 = Trim(ws.Range(mokc.Item("INF").Item("bomitem_qty").Item(1).Key & i))

If Len(str2) > 0 Then


If ws.Range(mokc.Item("INF").Item("bomitem_qty").Item(1).Key & i).Font.Strikethrough = True Or ws.Range(mokc.Item("INF").Item("bomitem_qty").Item(1).Key & i).EntireRow.Hidden = True Then
    If mokc.Item("ERROR") Is Nothing Then mokc.Add "ERROR", "ERROR"
    If mokc.Item("ERROR").Item("STRIKETHROUGH_OR_HIDDEN") Is Nothing Then mokc.Item("ERROR").Add "STRIKETHROUGH_OR_HIDDEN", "STRIKETHROUGH_OR_HIDDEN"
    mokc.Item("ERROR").Item("STRIKETHROUGH_OR_HIDDEN").Add CStr(i)
ElseIf mokc.Item("INF").Item("bom_tkid").Item(2).Key = ws.Range(mokc.Item("INF").Item("bomitem_tkid").Item(1).Key & i) Then
    If mokc.Item("ERROR") Is Nothing Then mokc.Add "ERROR", "ERROR"
    If mokc.Item("ERROR").Item("BOM_CONTAIN_ITSELF") Is Nothing Then mokc.Item("ERROR").Add "BOM_CONTAIN_ITSELF", "BOM_CONTAIN_ITSELF"
    mokc.Item("ERROR").Item("BOM_CONTAIN_ITSELF").Add CStr(i)
Else
    
    
    
    
'未填供应商的记录错误
str5 = Trim(ws.Range(mokc.Item("INF").Item("bomitem_manu").Item(1).Key & i))
If str5 = "" Then
str5 = ws.Range(mokc.Item("INF").Item("bomitem_tkid").Item(1).Key & i)
If Not (IS_BOM(str5)) Then
str5 = ws.Range(mokc.Item("INF").Item("bomitem_pos").Item(1).Key & i)
If str5 <> "0" Then
    If mokc.Item("ERROR") Is Nothing Then mokc.Add "ERROR", "ERROR"
    If mokc.Item("ERROR").Item("EMPTY_MANUFATURER") Is Nothing Then mokc.Item("ERROR").Add "EMPTY_MANUFATURER", "EMPTY_MANUFATURER"
    mokc.Item("ERROR").Item("EMPTY_MANUFATURER").Add CStr(i) & "_" & Trim(ws.Range(mokc.Item("INF").Item("bomitem_desc").Item(1).Key & i)) & "_" & Trim(ws.Range(mokc.Item("INF").Item("bomitem_tkid").Item(1).Key & i))
End If
End If
End If
'未填供应商的记录错误




j = j + 1
mokc.Item("BODY").Add CStr(i), CStr(i)
mokc.Item("BODY").Item(CStr(i)).Add str2, "bomitem_qty"


str3 = "bomitem_tkid"
str4 = ""
If Not (mokc.Item("INF").Item(str3) Is Nothing) Then str4 = ws.Range(mokc.Item("INF").Item(str3).Item(1).Key & i)
mokc.Item("BODY").Item(CStr(i)).Add str4, str3




str3 = "bomitem_pos"
str4 = ""
If Not (mokc.Item("INF").Item(str3) Is Nothing) Then str4 = ws.Range(mokc.Item("INF").Item(str3).Item(1).Key & i)
mokc.Item("BODY").Item(CStr(i)).Add str4, str3

str3 = "bomitem_type"
str4 = ""
If Not (mokc.Item("INF").Item(str3) Is Nothing) Then str4 = ws.Range(mokc.Item("INF").Item(str3).Item(1).Key & i)
mokc.Item("BODY").Item(CStr(i)).Add str4, str3

str3 = "bomitem_desc"
str4 = ""
If Not (mokc.Item("INF").Item(str3) Is Nothing) Then str4 = ws.Range(mokc.Item("INF").Item(str3).Item(1).Key & i)
mokc.Item("BODY").Item(CStr(i)).Add str4, str3

str3 = "bomitem_manu"
str4 = ""
If Not (mokc.Item("INF").Item(str3) Is Nothing) Then str4 = ws.Range(mokc.Item("INF").Item(str3).Item(1).Key & i)
mokc.Item("BODY").Item(CStr(i)).Add str4, str3

str3 = "bomitem_manu_ref"
str4 = ""
If Not (mokc.Item("INF").Item(str3) Is Nothing) Then str4 = ws.Range(mokc.Item("INF").Item(str3).Item(1).Key & i)
mokc.Item("BODY").Item(CStr(i)).Add str4, str3

str3 = "bomitem_norm"
str4 = ""
If Not (mokc.Item("INF").Item(str3) Is Nothing) Then str4 = ws.Range(mokc.Item("INF").Item(str3).Item(1).Key & i)
mokc.Item("BODY").Item(CStr(i)).Add str4, str3

str3 = "bomitem_dime"
str4 = ""
If Not (mokc.Item("INF").Item(str3) Is Nothing) Then str4 = ws.Range(mokc.Item("INF").Item(str3).Item(1).Key & i)
mokc.Item("BODY").Item(CStr(i)).Add str4, str3


str3 = "bomitem_unit"
str4 = ""
If Not (mokc.Item("INF").Item(str3) Is Nothing) Then str4 = ws.Range(mokc.Item("INF").Item(str3).Item(1).Key & i)
mokc.Item("BODY").Item(CStr(i)).Add str4, str3



str3 = "bomitem_mcbf"
str4 = ""
If Not (mokc.Item("INF").Item(str3) Is Nothing) Then
str4 = ws.Range(mokc.Item("INF").Item(str3).Item(1).Key & i)
mokc.Item("BODY").Item(CStr(i)).Add str4, str3
End If

str3 = "bomitem_pc"
str4 = ""
If Not (mokc.Item("INF").Item(str3) Is Nothing) Then
str4 = ws.Range(mokc.Item("INF").Item(str3).Item(1).Key & i)
mokc.Item("BODY").Item(CStr(i)).Add str4, str3
End If

str3 = "bomitem_postype"
str4 = ""
If Not (mokc.Item("INF").Item(str3) Is Nothing) Then
str4 = ws.Range(mokc.Item("INF").Item(str3).Item(1).Key & i)
mokc.Item("BODY").Item(CStr(i)).Add str4, str3
End If


str3 = "bomitem_ver"
str4 = ""
If Not (mokc.Item("INF").Item(str3) Is Nothing) Then
str4 = ws.Range(mokc.Item("INF").Item(str3).Item(1).Key & i)
mokc.Item("BODY").Item(CStr(i)).Add str4, str3
End If



End If


End If








Next

End If
'给可以识别的ＢＯＭ赋　值











Read_BOM = True

End Function


Private Function Get_FlnTkid(fln As String) As String
'本函数用于提取字符串首的TKID

If InStr(fln, "\") > 0 Then
fln = Right(fln, Len(fln) - InStrRev(fln, "\"))
End If

If (fln Like "?.?????.???.??.??*") And Not (fln Like "?.?????.???.TZ.??*") And Not (fln Like "?.?????.???.EN.??*") And Not (fln Like "?.?????.???.EH.??*") Then
Get_FlnTkid = Left(fln, 17)
Else
Get_FlnTkid = "UNKNOWN"
End If




If Left(Get_FlnTkid, 2) = "d." Then Get_FlnTkid = "D." & Right(Get_FlnTkid, Len(Get_FlnTkid) - 2)


End Function

Private Function IS_BOM(str1 As String) As Boolean
IS_BOM = False
If str1 Like "?.?????.???.ST.??" Then IS_BOM = True
End Function

Public Function get_bom_Body_Count(TKID_BOM As String, Optional KEY_FLN_OR_HASH As String = "") As Integer

On Error GoTo ERRORHAND
If mokc.Item(TKID_BOM) Is Nothing Then
get_bom_Body_Count = 0

ElseIf mokc.Item(TKID_BOM).Item(KEY_FLN_OR_HASH) Is Nothing Then
get_bom_Body_Count = mokc.Item(TKID_BOM).Item(1).Item("BODY").Count
Else
get_bom_Body_Count = mokc.Item(TKID_BOM).Item(KEY_FLN_OR_HASH).Item("BODY").Count
End If
Exit Function
ERRORHAND:
get_bom_Body_Count = 0
End Function


Public Function Bom_Check_OK(TKID_BOM As String, Optional KEY_FLN_OR_HASH As String = "") As Boolean

On Error GoTo ERRORHAND
Bom_Check_OK = True

If mokc.Item(TKID_BOM) Is Nothing Then
Bom_Check_OK = False


ElseIf mokc.Item(TKID_BOM).Item(KEY_FLN_OR_HASH) Is Nothing Then

If mokc.Item(TKID_BOM).Item(1).Count = 3 Then
Bom_Check_OK = False
End If

Else

If mokc.Item(TKID_BOM).Item(KEY_FLN_OR_HASH).Count = 3 Then
Bom_Check_OK = False
End If

End If
Exit Function
ERRORHAND:
Bom_Check_OK = False
End Function

Private Function open_wb(ByRef wb As Workbook, ByVal FLFP As String) As Boolean
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
Set wb = Workbooks.Open(flp & fln, 0)

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

Public Function Bom_Error(TKID_BOM As String, Optional KEY_FLN_OR_HASH As String = "") As String

'BOM_NOT_EXIST
'UNEXPECT_ERROR
'...


On Error GoTo ERRORHAND
Bom_Error = ""



If mokc.Item(TKID_BOM) Is Nothing Then
Bom_Error = "BOM_NOT_EXIST"

ElseIf mokc.Item(TKID_BOM).Item(KEY_FLN_OR_HASH) Is Nothing Then

If mokc.Item(TKID_BOM).Item(1).Count = 3 Then
Bom_Error = mokc.Item(TKID_BOM).Item(1).Item(3).To_str("", Chr(10))


End If

Else

If mokc.Item(TKID_BOM).Item(KEY_FLN_OR_HASH).Count = 3 Then
Bom_Error = mokc.Item(TKID_BOM).Item(KEY_FLN_OR_HASH).Item(3).To_str("", Chr(10))

End If

End If
Exit Function
ERRORHAND:
Bom_Error = "UNEXPECT_ERROR"
End Function

