Attribute VB_Name = "FileFolderProcess"
Option Explicit

'----------------------------------------------------------------------------------------------------
' ファイルが存在するかどうかを確認する
'
' 引数
'   filepath    : ファイルパス
' 戻り値
'   True    : 存在する
'   False   : 存在しない
'----------------------------------------------------------------------------------------------------
Function IsFileExists(filePath As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    IsFileExists = fso.FileExists(filePath)
End Function

'----------------------------------------------------------------------------------------------------
' ファイルパスからファイル名だけ取り出す
'
' 引数
'   filepath    : ファイルパス
' 戻り値
'   ファイル名
'----------------------------------------------------------------------------------------------------
Function GetFileName(ByVal filePath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetFileName = fso.GetFileName(filePath)
End Function

'----------------------------------------------------------------------------------------------------
' ファイルパスからフォルダー名を取り出す
'
' 引数
'   filepath    : ファイルパス
' 戻り値
'   フォルダー名
'----------------------------------------------------------------------------------------------------
Function GetParentFolder(ByVal filePath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetParentFolder = fso.GetParentFolderName(filePath)
End Function

'----------------------------------------------------------------------------------------------------
' ファイルパスから拡張子を取り除いたファイル名だけ取り出す
'
' 引数
'   filepath    : ファイルパス
' 戻り値
'   ファイル名（拡張子なし）
'----------------------------------------------------------------------------------------------------
Function GetBaseName(ByVal filePath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetBaseName = fso.GetBaseName(filePath)
End Function

'----------------------------------------------------------------------------------------------------
' ファイルパスから拡張子を取り出す
'
' 引数
'   filepath    : ファイルパス
' 戻り値
'   拡張子(ピリオドはつかない)
'----------------------------------------------------------------------------------------------------
Function GetExtention(ByVal filePath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetExtention = fso.GetExtensionName(filePath)
End Function
