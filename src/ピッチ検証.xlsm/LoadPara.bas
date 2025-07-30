Attribute VB_Name = "LoadPara"
Option Explicit

Public Sub ParameterLoad()
    Dim result As Variant
    Dim filters As Variant
    filters = Array( _
        Array("LOT�t�@�C��", "*.LOT"), _
        Array("WDT�t�@�C��", "*.WDT") _
    )
    Dim filePath As String
    
    result = ShowFileDialogWithFilters("�p�����[�^�̓Ǎ�", False, filters)

    If Not IsArray(result) Then
        If VarType(result) = vbString Then
            filePath = CStr(result)
        Else
            Exit Sub
        End If
    Else
        Exit Sub
    End If

    If Not IsFileExists(filePath) Then
        MsgBox "�t�@�C����������܂���", vbExclamation, "���[�j���O"
        Exit Sub
    End If
    
    Dim ext As String
    ext = UCase(GetExtention(filePath))
    
    Sheet1.Cells(1, "E") = GetFileName(filePath)
    
    If ext = "LOT" Then
        LotFileLoad filePath
    Else
        WDTFileLoad filePath
    End If

End Sub

' LOT�t�@�C���̓Ǎ��ƃf�[�^�̕\��
Public Sub LotFileLoad(ByVal filePath As String)
    On Error GoTo ErrHandler
    
    Dim ini As INIFile
    Dim wdcPara As Collection
    Dim bWDC As Boolean
    Dim stepNum As Long
    Dim i As Long
    Dim alpha As Integer
    Dim beta As Integer
    
    Set wdcPara = New Collection
    
    If Not IsFileExists(filePath) Then
        MsgBox "�t�@�C��������܂���" & vbCrLf & filePath, vbExclamation, "���[�j���O"
        Exit Sub
    End If
    
    Set ini = New INIFile
    
    If ini.OpenFile(filePath) Then
        bWDC = ini.ReadValueAsBoolean("HEAD", "WDC")
        If Not bWDC Then
            MsgBox "LOT�t�@�C����WDC�p�����[�^���܂܂�Ă��܂���B", vbExclamation, "���[�j���O"
            Exit Sub
        End If
        
        If LOTFastRead(0, filePath, wdcPara) Then
            ClearFromRow Sheet1, 6
            ' TR Stroke
            Sheet1.Cells(1, "C") = CDbl(wdcPara(993)) / 10#
            ' Step��
            stepNum = wdcPara(100)
            Sheet1.Cells(2, "C") = stepNum
            ' ���C���h��e�[�u��
            For i = 0 To stepNum - 1
                Sheet1.Cells(6 + i, "A") = i + 1    ' Step No
                Sheet1.Cells(6 + i, "B") = CDbl(wdcPara(305 + ((i Mod 5) * 3) + (Int(i / 5) * 16))) / 10#  ' ���a
                alpha = wdcPara(306 + ((i Mod 5) * 3) + (Int(i / 5) * 16))
                beta = wdcPara(307 + ((i Mod 5) * 3) + (Int(i / 5) * 16))
                Sheet1.Cells(6 + i, "C") = CDbl(alpha) / 10# + CDbl(beta) / 100000# ' ���C���h��
            Next i
        Else
            MsgBox "WDC�p�����[�^�̃��[�h�Ɏ��s���܂����B", vbExclamation, "���[�j���O"
        End If
        
    Else
        MsgBox "LOT�t�@�C���̓Ǎ��Ɏ��s���܂���", vbExclamation, "�G���["
        Exit Sub
    End If
    
    Sheet1.CommandButton1_Click

    Exit Sub

ErrHandler:
    MsgBox "LOT�t�@�C���̓Ǎ��Ɏ��s���܂���", vbExclamation, "�G���["
End Sub

' WDT�t�@�C���̓Ǎ��ƃf�[�^�\��
Public Sub WDTFileLoad(ByVal filePath As String)
    On Error GoTo ErrHandler
    
    Dim wdcPara As Collection
    Dim stepNum As Long
    Dim i As Long
    Dim alpha As Integer
    Dim beta As Integer
    
    Set wdcPara = New Collection
    
    If WDTFileRead(filePath, wdcPara) Then
        ClearFromRow Sheet1, 6
        ' TR Stroke
        Sheet1.Cells(1, "C") = CDbl(wdcPara(993)) / 10#
        ' Step��
        stepNum = wdcPara(100)
        Sheet1.Cells(2, "C") = stepNum
        ' ���C���h��e�[�u��
        For i = 0 To stepNum - 1
            Sheet1.Cells(6 + i, "A") = i + 1    ' Step No
            Sheet1.Cells(6 + i, "B") = CDbl(wdcPara(305 + ((i Mod 5) * 3) + (Int(i / 5) * 16))) / 10#  ' ���a
            alpha = wdcPara(306 + ((i Mod 5) * 3) + (Int(i / 5) * 16))
            beta = wdcPara(307 + ((i Mod 5) * 3) + (Int(i / 5) * 16))
            Sheet1.Cells(6 + i, "C") = CDbl(alpha) / 10# + CDbl(beta) / 100000# ' ���C���h��
        Next i
    Else
        MsgBox "WDC�p�����[�^�̃��[�h�Ɏ��s���܂����B", vbExclamation, "���[�j���O"
    End If

    Sheet1.CommandButton1_Click

    Exit Sub

ErrHandler:
    MsgBox "WDC�p�����[�^�t�@�C���̓Ǎ��Ɏ��s���܂���", vbExclamation, "�G���["
End Sub

' LOT�t�@�C���������ɓǂݏo��
' �p�����[�^���ԍ����ɕ���ł��邱�Ƃ�O��Ƃ��Ă̓ǂݏo��
Private Function LOTFastRead(kind As Integer, ByVal filePath As String, ByRef values As Collection) As Boolean
    On Error GoTo ErrHandler
    
    Dim fileNum
    Dim s As String
    Dim items() As String
    Dim items2() As String
    Dim no As Integer
    Dim secStr As String
    Dim found As Boolean
    Dim maxParaNum As Integer
    Dim val As Integer
    
    If Not IsFileExists(filePath) Then
        LOTFastRead = False
        Exit Function
    End If
    
    fileNum = FreeFile
    
    If kind = 0 Then
        secStr = "[WDC]"
        maxParaNum = 2048
    ElseIf kind = 1 Then
        secStr = "[GDC]"
        maxParaNum = 2048
    Else
        secStr = "[AUX]"
        maxParaNum = 512
    End If
    
    Open filePath For Input As #fileNum
    
    found = False
    Do Until EOF(fileNum)
        Line Input #fileNum, s
        s = Trim(s)
        If Not found Then
            If InStr(1, s, secStr, vbTextCompare) >= 1 Then
                found = True
            End If
        Else
            If Left(s, 1) <> ";" Or s = "" Then         ' �R�����g�s���s��ǂݔ�΂�
                items = Split(s, "=")                   ' key=****,*****�Ȃ̂�Key�Ɛݒ�𕪂���
                If UBound(items) >= 1 Then
                    no = CInt(items(0))                 ' �p�����[�^�ԍ�(0�`)
                    If no < maxParaNum Then
                        items2 = Split(items(1), ",")
                        If UBound(items2) >= 1 Then
                            val = CInt(items2(0))
                            values.Add val
                        End If
                    End If
                    
                    If (no + 1) = maxParaNum Then
                        Exit Do
                    End If
                End If
            End If
        End If
    Loop
    
    Close #fileNum
    
    If (no + 1) = maxParaNum Then
        LOTFastRead = True
    Else
        LOTFastRead = False
    End If
    
    Exit Function
    
ErrHandler:

    LOTFastRead = False
    
End Function

' WDT�t�@�C���̓Ǎ��ƃV�[�g�ւ̕\��
Private Function WDTFileRead(ByVal filePath As String, ByRef values As Collection) As Boolean
    On Error GoTo ErrHandler
    Dim fileNum
    Dim s As String
    Dim items() As String
    Dim para As Integer
    
    If Not IsFileExists(filePath) Then
        WDTFileRead = False
        Exit Function
    End If
    
    num = 0
    fileNum = FreeFile
    
    Open filePath For Input As #fileNum
    
    Do Until EOF(fileNum)
        Line Input #fileNum, s
        s = Trim(s)
        If Left(s, 1) <> ";" Or s = "" Then         ' �R�����g�s���s��ǂݔ�΂�
            items = Split(s, ",")
            If UBound(items) >= 0 Then
                If IsNumeric(Trim(items(0))) Then
                    para = CInt(Trim(items(0)))
                    values.Add para
                End If
            End If
        End If
    Loop
    
    Close #fileNum
    
    If values.Count = 2048 Then
        WDTFileRead = True
    Else
        WDTFileRead = False
    End If
    
    Exit Function
    
ErrHandler:

    WDTFileRead = False
    
End Function
