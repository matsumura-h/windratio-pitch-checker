Attribute VB_Name = "FileFolderProcess"
Option Explicit

'----------------------------------------------------------------------------------------------------
' �t�@�C�������݂��邩�ǂ������m�F����
'
' ����
'   filepath    : �t�@�C���p�X
' �߂�l
'   True    : ���݂���
'   False   : ���݂��Ȃ�
'----------------------------------------------------------------------------------------------------
Function IsFileExists(filePath As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    IsFileExists = fso.FileExists(filePath)
End Function

'----------------------------------------------------------------------------------------------------
' �t�@�C���p�X����t�@�C�����������o��
'
' ����
'   filepath    : �t�@�C���p�X
' �߂�l
'   �t�@�C����
'----------------------------------------------------------------------------------------------------
Function GetFileName(ByVal filePath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetFileName = fso.GetFileName(filePath)
End Function

'----------------------------------------------------------------------------------------------------
' �t�@�C���p�X����t�H���_�[�������o��
'
' ����
'   filepath    : �t�@�C���p�X
' �߂�l
'   �t�H���_�[��
'----------------------------------------------------------------------------------------------------
Function GetParentFolder(ByVal filePath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetParentFolder = fso.GetParentFolderName(filePath)
End Function

'----------------------------------------------------------------------------------------------------
' �t�@�C���p�X����g���q����菜�����t�@�C�����������o��
'
' ����
'   filepath    : �t�@�C���p�X
' �߂�l
'   �t�@�C�����i�g���q�Ȃ��j
'----------------------------------------------------------------------------------------------------
Function GetBaseName(ByVal filePath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetBaseName = fso.GetBaseName(filePath)
End Function

'----------------------------------------------------------------------------------------------------
' �t�@�C���p�X����g���q�����o��
'
' ����
'   filepath    : �t�@�C���p�X
' �߂�l
'   �g���q(�s���I�h�͂��Ȃ�)
'----------------------------------------------------------------------------------------------------
Function GetExtention(ByVal filePath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetExtention = fso.GetExtensionName(filePath)
End Function
