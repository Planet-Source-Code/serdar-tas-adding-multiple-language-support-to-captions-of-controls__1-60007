Attribute VB_Name = "MDBConnection"
Option Explicit
Global Language As Integer
Public Rs As New ADODB.Recordset
Public Cn As New ADODB.Connection
Public Cmd As New ADODB.Command
Private SQL As String, CTL As Control, i As Integer
Public Sub Connect(Path As String)
    'On Error GoTo Hell
    Cn.ConnectionString = "Provider=Microsoft.jet.OLEDB.4.0;Data Source=" & Path & ";Persist Security Info=False"
    Cn.CursorLocation = adUseClient
    Cn.Mode = adModeReadWrite
    Cn.Open
    Exit Sub
Hell:
    MsgBox Err.Number & " : " & Err.Description
End Sub

Public Function LanguageID(FileName As String) As Integer
On Error Resume Next
LanguageID = 1
    Open FileName For Input As #1
        Line Input #1, SQL
            LanguageID = Val(SQL)
    Close #1
End Function
Public Sub SaveLanguageID(FileName As String, LanguageID As Integer)
On Error Resume Next
    Open FileName For Output As #1
    Print #1, Str(LanguageID)
    Close #1
End Sub
Public Sub SelectLanguage(FormName As Form, LanguageID As Integer)
On Error Resume Next
    For Each CTL In FormName.Controls
        CTL.Caption = ReadCaption(Val(CTL.Tag), LanguageID)
    Next
End Sub

Private Function ReadCaption(FieldID As Integer, LanguageID As Integer) As String
    Set Rs = New ADODB.Recordset
    Rs.Open "Select FieldName from Tbl_ControlNames where FieldID = " & FieldID & " and LanguageId = " & LanguageID, Cn, adOpenDynamic, adLockReadOnly
    If Rs.RecordCount <> 0 Then
        ReadCaption = Rs.Fields(0)
    Else
        ReadCaption = Null
    End If
End Function

Public Sub FillCombo(Combo As ComboBox, SQL As String)
    Combo.Clear
    Combo.AddItem ""
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Cn, adOpenKeyset, adLockReadOnly
    If Not Rs.RecordCount = 0 Then
        For i = 1 To Rs.RecordCount
            Combo.AddItem Rs(1)
            Combo.ItemData(Combo.ListCount - 1) = CInt(Rs(0))
            Rs.MoveNext
        Next i
        Combo.ListIndex = 0
    End If
    Set Rs = Nothing
End Sub
Public Sub PopUp(MessageId As Integer, LanguageID As Integer, PopUpType As VbMsgBoxStyle)
    Set Rs = New ADODB.Recordset
    Rs.Open "SELECT Alert, Title From Cnst_Alerts Where AlertID = " & MessageId & " And LanguageID = " & LanguageID, Cn, adOpenDynamic, adLockReadOnly
    MsgBox Rs.Fields(0), PopUpType, Rs.Fields(1)
    Set Rs = Nothing
End Sub
Public Sub GridHeaders(GridName As MSFlexGrid, HeaderID As Integer, LanguageID As Integer)
    Set Rs = New ADODB.Recordset
    Rs.Open "Select Column, Header From Tbl_GridHeaders Where HeaderID = " & HeaderID & " And LanguageID = " & LanguageID, Cn, adOpenDynamic, adLockReadOnly
    For i = 0 To Rs.RecordCount - 1
        GridName.TextMatrix(0, Val(Rs.Fields(0))) = Rs.Fields(1)
        Rs.MoveNext
    Next i
End Sub
