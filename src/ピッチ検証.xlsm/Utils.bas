Attribute VB_Name = "Utils"
Option Explicit

'-----------------------------------------------------
'   GCD(Greatest Common Divisor) �ő����
'
'   ���[�N���b�h�̌ݏ��@�ŋ��߂�(�ċA��)
'-----------------------------------------------------
Public Function gcd(ByVal m As Long, ByVal N As Long) As Long
Dim r As Integer

    r = m Mod N
    If r = 0 Then
        gcd = N
    Else
        gcd = gcd(N, r)
    End If
    
End Function

'-----------------------------------------------------
'   LCM(Least Common Multiple) �ŏ����{��
'-----------------------------------------------------
Public Function lcm(ByVal m As Long, ByVal N As Long) As Long

    If m = 0 Or N = 0 Then
        lcm = 0
    Else
        lcm = (m / gcd(m, N)) * N
    End If

End Function

'����������P���A������
Function ContFrac(x#, ByVal N%, B&()) As Integer
    On Error Resume Next
    Dim i%, XX#
    
    For i = 0 To N
        B(i) = -1
    Next
    
    XX = x
    B&(0) = Fix(XX)
    For i% = 1 To N%
        If (XX - CDbl(B&(i% - 1))) < 0.000001 Then
            Exit For
        End If
        XX = 1# / (XX - CDbl(B&(i% - 1)))
        B&(i) = Fix(XX)
        If B&(i) >= 1000 Then
            B&(i) = -1
            Exit For
        End If
    Next i%
    
    '�Ōオ1�̏ꍇ��1/1�ɂȂ�̂ł��̑O�̒l��+1���Ď��������炷
'    If i < N And B&(i - 1) = 1 Then
'        B&(i - 2) = B&(i - 2) + 1
'        B&(i - 1) = -1
'        i = i - 1
'    End If
    
    ContFrac = i
End Function

Function deg2rad(deg As Double) As Double
    deg2rad = deg * PI# / 180#
End Function

'�A�����𕪐���
Sub FracConv(ByVal N%, B&(), F&, G&)
    On Error Resume Next
    Dim i%, temp&, d&
    
'    For i = 0 To N
'        If (B(i) = -1) Then
'            Exit For
'        End If
'    Next
'    N = i - 1
    
    F& = B&(N%)
    G& = 1
    For i% = N% - 1 To 0 Step -1
        temp& = B&(i%) * F& + G&
        G& = F&
        F& = temp&
        d& = gcd(F&, G&)
        F& = F& \ d&
        G& = G& \ d&
    Next i%
End Sub

'----------------------------------------------------------------------------------------------------
' �t�@�C���I���_�C�A���O
'
' ����
'   Title       : �_�C�A���O�^�C�g��
'   initFolder  : �����\���t�H���_�[
'   fileType    : �t�@�C�����(��:Excel�t�@�C��)
'   filter      : �g���q�t�B���^�[(��:*.xlsx)
'----------------------------------------------------------------------------------------------------
Function OpenFileDialog(ByVal Title As String, ByVal initFolder As String, ByVal fileType As String, ByVal filter As String) As String
    Dim fd As FileDialog
    Dim filePath As String

    ' FileDialog �I�u�W�F�N�g�擾
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = Title
        .InitialFileName = initFolder
        .filters.Clear
        .filters.Add fileType, filter
        .AllowMultiSelect = False

        ' �_�C�A���O�\��
        If .Show = -1 Then
            ' �I�����ꂽ�t�@�C�����擾
            OpenFileDialog = .SelectedItems(1)
        Else
            OpenFileDialog = ""
        End If
    End With

    Set fd = Nothing

End Function

'----------------------------------------------------------------------------------------------------
' �t�@�C���I���_�C�A���O
'
' ����
'   Title               : �_�C�A���O�^�C�g��
'   AllowMultiSelect    : True=�����t�@�C���I���AFalse=�P��t�@�C��
'   filters             : �g���q�t�B���^�[�̔z��
'----------------------------------------------------------------------------------------------------
Function ShowFileDialogWithFilters( _
    Optional Title As String = "�t�@�C����I�����Ă�������", _
    Optional AllowMultiSelect As Boolean = False, _
    Optional filters As Variant = Empty _
) As Variant
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = Title
        .AllowMultiSelect = AllowMultiSelect
        .filters.Clear
        
        ' �t�B���^�[���w�肳��Ă���Βǉ�
        If Not IsEmpty(filters) Then
            Dim i As Long
            For i = LBound(filters) To UBound(filters)
                ' �e�t�B���^�[�͔z��`��: Array("�\����", "*.ext1;*.ext2")
                If IsArray(filters(i)) And UBound(filters(i)) = 1 Then
                    .filters.Add filters(i)(0), filters(i)(1)
                End If
            Next i
        End If
        
        ' �_�C�A���O�\��
        If .Show = -1 Then
            If AllowMultiSelect Then
                ShowFileDialogWithFilters = .SelectedItems
            Else
                ShowFileDialogWithFilters = .SelectedItems(1)
            End If
        Else
            ShowFileDialogWithFilters = False
        End If
    End With
End Function


'�w�肵���V�[�g�̎w�肵���s����Ō�܂ŃN���A����
Sub ClearFromRow(ws As Worksheet, startRow As Long)
    Dim lastRow As Long
    
    ' �ŏI�s���擾
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    
    ' �w�肵���s����ŏI�s�܂ł��N���A
    If lastRow >= startRow Then
        ws.Range(ws.Cells(startRow, 1), ws.Cells(lastRow, ws.Columns.Count)).ClearContents
    End If
End Sub

