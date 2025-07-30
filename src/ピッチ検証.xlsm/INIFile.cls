VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "INIFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' INIFile Class Module
' エンコーディング指定可能なINIファイルアクセスクラス
'==============================================================================

Option Explicit

Private m_filePath As String
Private m_encoding As String
Private m_lines As Collection  ' 全ての行を保持
Private m_isLoaded As Boolean

Private Const DEFAULT_ENCODING As String = "Shift_JIS"  ' "UTF-8", "GB2312"など

'==============================================================================
' プロパティ
'==============================================================================

Public Property Get filePath() As String
    filePath = m_filePath
End Property

Public Property Get Encoding() As String
    Encoding = m_encoding
End Property

Public Property Get IsLoaded() As Boolean
    IsLoaded = m_isLoaded
End Property

'==============================================================================
' パブリックメソッド
'==============================================================================

' INIファイルを開く
Public Function OpenFile(filePath As String, Optional Encoding As String = DEFAULT_ENCODING) As Boolean
    On Error GoTo ErrorHandler
    
    m_filePath = filePath
    m_encoding = Encoding
    m_isLoaded = False
    
    Set m_lines = New Collection
    
    If Dir(filePath) <> "" Then
        If LoadFromFile() Then
            m_isLoaded = True
            OpenFile = True
        End If
    Else
        ' ファイルが存在しない場合は新規作成として扱う
        m_isLoaded = True
        OpenFile = True
    End If
    
    Exit Function
    
ErrorHandler:
    OpenFile = False
End Function

' INIファイルを保存
Public Function SaveFile(Optional filePath As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    If filePath <> "" Then
        m_filePath = filePath
    End If
    
    If m_filePath = "" Then
        SaveFile = False
        Exit Function
    End If
    
    SaveFile = SaveToFile()
    
    Exit Function
    
ErrorHandler:
    SaveFile = False
End Function

' 値を読み取り
Public Function ReadValue(sectionName As String, keyName As String, Optional defaultValue As String = "") As String
    On Error GoTo ErrorHandler
    
    If Not m_isLoaded Then
        ReadValue = defaultValue
        Exit Function
    End If
    
    Dim targetSection As String
    targetSection = "[" & sectionName & "]"
    Dim inSection As Boolean
    inSection = False
    
    Dim i As Long
    For i = 1 To m_lines.Count
        Dim line As String
        line = Trim(m_lines.Item(i))
        
        ' セクション行の処理
        If Left(line, 1) = "[" Then
            Dim closeBracketPos As Long
            closeBracketPos = InStr(line, "]")
            If closeBracketPos > 0 Then
                Dim sectionHeader As String
                sectionHeader = Left(line, closeBracketPos)
                If UCase(sectionHeader) = UCase(targetSection) Then
                    inSection = True
                Else
                    ' 別のセクションに入った場合、目的のセクションを抜ける
                    If inSection Then
                        Exit For
                    End If
                    inSection = False
                End If
            End If
        ElseIf inSection And line <> "" And Left(line, 1) <> ";" And InStr(line, "=") > 0 Then
            ' キー=値の行
            Dim semicolonPos As Long
            semicolonPos = InStr(line, ";")
            
            Dim keyValuePart As String
            If semicolonPos > 0 Then
                keyValuePart = Trim(Left(line, semicolonPos - 1))
            Else
                keyValuePart = line
            End If
            
            Dim parts() As String
            parts = Split(keyValuePart, "=", 2)
            
            If UBound(parts) >= 1 And UCase(Trim(parts(0))) = UCase(keyName) Then
                ReadValue = Trim(parts(1))
                Exit Function
            End If
        End If
    Next i
    
    ReadValue = defaultValue
    Exit Function
    
ErrorHandler:
    ReadValue = defaultValue
End Function

' 整数値を読み取り
Public Function ReadValueAsInteger(sectionName As String, keyName As String, Optional defaultValue As Long = 0) As Long
    On Error GoTo ErrorHandler
    
    Dim strValue As String
    strValue = ReadValue(sectionName, keyName, CStr(defaultValue))
    
    ' 数値への変換を試行
    If IsNumeric(strValue) Then
        ReadValueAsInteger = CLng(strValue)
    Else
        ReadValueAsInteger = defaultValue
    End If
    
    Exit Function
    
ErrorHandler:
    ReadValueAsInteger = defaultValue
End Function

' 浮動小数点値を読み取り
Public Function ReadValueAsDouble(sectionName As String, keyName As String, Optional defaultValue As Double = 0#) As Double
    On Error GoTo ErrorHandler
    
    Dim strValue As String
    strValue = ReadValue(sectionName, keyName, CStr(defaultValue))
    
    ' 数値への変換を試行
    If IsNumeric(strValue) Then
        ReadValueAsDouble = CDbl(strValue)
    Else
        ReadValueAsDouble = defaultValue
    End If
    
    Exit Function
    
ErrorHandler:
    ReadValueAsDouble = defaultValue
End Function

' Boolean値を読み取り
Public Function ReadValueAsBoolean(sectionName As String, keyName As String, Optional defaultValue As Boolean = False) As Boolean
    On Error GoTo ErrorHandler
    
    Dim strValue As String
    strValue = UCase(Trim(ReadValue(sectionName, keyName, IIf(defaultValue, "True", "False"))))
    
    ' 様々なBoolean表現に対応
    Select Case strValue
        Case "TRUE", "YES", "ON", "1", "ENABLED", "ENABLE"
            ReadValueAsBoolean = True
        Case "FALSE", "NO", "OFF", "0", "DISABLED", "DISABLE"
            ReadValueAsBoolean = False
        Case Else
            ReadValueAsBoolean = defaultValue
    End Select
    
    Exit Function
    
ErrorHandler:
    ReadValueAsBoolean = defaultValue
End Function

' 値を書き込み
Public Function WriteValue(sectionName As String, keyName As String, value As String) As Boolean
    On Error GoTo ErrorHandler
    
    If Not m_isLoaded Then
        WriteValue = False
        Exit Function
    End If
    
    Dim targetSection As String
    targetSection = "[" & sectionName & "]"
    Dim inSection As Boolean
    inSection = False
    Dim sectionStartLine As Long
    sectionStartLine = 0
    Dim keyLineIndex As Long
    keyLineIndex = 0
    
    ' セクションを探し、既存のキーがあるかチェック
    Dim i As Long
    For i = 1 To m_lines.Count
        Dim line As String
        line = Trim(m_lines.Item(i))
        
        ' セクション行の処理
        If Left(line, 1) = "[" Then
            Dim closeBracketPos As Long
            closeBracketPos = InStr(line, "]")
            If closeBracketPos > 0 Then
                Dim sectionHeader As String
                sectionHeader = Left(line, closeBracketPos)
                If UCase(sectionHeader) = UCase(targetSection) Then
                    inSection = True
                    sectionStartLine = i
                Else
                    If inSection Then
                        ' 次のセクションに入ったので終了
                        Exit For
                    End If
                    inSection = False
                End If
            End If
        ElseIf inSection And line <> "" And Left(line, 1) <> ";" And InStr(line, "=") > 0 Then
            ' キー=値の行
            Dim semicolonPos As Long
            semicolonPos = InStr(line, ";")
            
            Dim keyValuePart As String
            If semicolonPos > 0 Then
                keyValuePart = Trim(Left(line, semicolonPos - 1))
            Else
                keyValuePart = line
            End If
            
            Dim parts() As String
            parts = Split(keyValuePart, "=", 2)
            
            If UBound(parts) >= 1 And UCase(Trim(parts(0))) = UCase(keyName) Then
                ' 既存のキーを更新（コメントを保持）
                Dim newLine As String
                If semicolonPos > 0 Then
                    newLine = keyName & "=" & value & " " & Mid(m_lines.Item(i), semicolonPos)
                Else
                    newLine = keyName & "=" & value
                End If
                
                ' 新しいCollectionを作成して行を置き換え
                Dim newLines As Collection
                Set newLines = New Collection
                
                Dim j As Long
                For j = 1 To m_lines.Count
                    If j = i Then
                        newLines.Add newLine
                    Else
                        newLines.Add m_lines.Item(j)
                    End If
                Next j
                
                Set m_lines = newLines
                WriteValue = True
                Exit Function
            End If
        End If
    Next i
    
    ' セクションが見つからない場合は作成
    If sectionStartLine = 0 Then
        ' ファイルが空でない場合のみ空行を追加
        If m_lines.Count > 0 Then
            m_lines.Add ""
        End If
        m_lines.Add targetSection
        m_lines.Add keyName & "=" & value
        WriteValue = True
        Exit Function
    End If
    
    ' セクション内に新しいキーを追加
    ' セクションの最後（次のセクションの直前または空行の直前）に追加
    Dim insertPos As Long
    
    ' セクションの最後の位置を特定
    If i <= m_lines.Count Then
        ' 次のセクションまたはファイル末尾の位置
        insertPos = i - 1
    Else
        ' ファイル末尾
        insertPos = m_lines.Count
    End If
    
    ' 空行がある場合は空行の前に挿入するため、空行を探して位置を調整
    Do While insertPos > sectionStartLine And insertPos <= m_lines.Count
        If Trim(m_lines.Item(insertPos)) = "" Then
            insertPos = insertPos - 1
        Else
            Exit Do
        End If
    Loop
    
    ' insertPosの後に挿入（afterパラメータを使用）
    If insertPos >= m_lines.Count Then
        ' ファイル末尾の場合
        m_lines.Add keyName & "=" & value
    Else
        ' 指定位置の後に挿入
        m_lines.Add keyName & "=" & value, , , insertPos
    End If
    WriteValue = True
    
    Exit Function
    
ErrorHandler:
    WriteValue = False
End Function

' 整数値を書き込み
Public Function WriteValueAsInteger(sectionName As String, keyName As String, value As Long) As Boolean
    WriteValueAsInteger = WriteValue(sectionName, keyName, CStr(value))
End Function

' 浮動小数点値を書き込み
Public Function WriteValueAsDouble(sectionName As String, keyName As String, value As Double) As Boolean
    WriteValueAsDouble = WriteValue(sectionName, keyName, CStr(value))
End Function

' Boolean値を書き込み
Public Function WriteValueAsBoolean(sectionName As String, keyName As String, value As Boolean) As Boolean
    WriteValueAsBoolean = WriteValue(sectionName, keyName, IIf(value, "True", "False"))
End Function

' セクションを削除
Public Function RemoveSection(sectionName As String) As Boolean
    On Error GoTo ErrorHandler
    
    If Not m_isLoaded Then
        RemoveSection = False
        Exit Function
    End If
    
    Dim targetSection As String
    targetSection = "[" & sectionName & "]"
    Dim inSection As Boolean
    inSection = False
    Dim linesToRemove As Collection
    Set linesToRemove = New Collection
    
    ' 削除する行を特定
    Dim i As Long
    For i = 1 To m_lines.Count
        Dim line As String
        line = Trim(m_lines.Item(i))
        
        If Left(line, 1) = "[" Then
            Dim closeBracketPos As Long
            closeBracketPos = InStr(line, "]")
            If closeBracketPos > 0 Then
                Dim sectionHeader As String
                sectionHeader = Left(line, closeBracketPos)
                If UCase(sectionHeader) = UCase(targetSection) Then
                    inSection = True
                    linesToRemove.Add i
                Else
                    If inSection Then
                        ' 次のセクションに入ったので終了
                        Exit For
                    End If
                End If
            End If
        ElseIf inSection Then
            linesToRemove.Add i
        End If
    Next i
    
    ' 後ろから削除（インデックスがずれないように）
    For i = linesToRemove.Count To 1 Step -1
        m_lines.Remove linesToRemove.Item(i)
    Next i
    
    RemoveSection = True
    
    Exit Function
    
ErrorHandler:
    RemoveSection = False
End Function

' キーを削除
Public Function RemoveKey(sectionName As String, keyName As String) As Boolean
    On Error GoTo ErrorHandler
    
    If Not m_isLoaded Then
        RemoveKey = False
        Exit Function
    End If
    
    Dim targetSection As String
    targetSection = "[" & sectionName & "]"
    Dim inSection As Boolean
    inSection = False
    Dim lineToRemove As Long
    lineToRemove = 0
    
    ' 削除対象のキーを探す
    Dim i As Long
    For i = 1 To m_lines.Count
        Dim line As String
        line = Trim(m_lines.Item(i))
        
        ' セクション行の処理
        If Left(line, 1) = "[" Then
            Dim closeBracketPos As Long
            closeBracketPos = InStr(line, "]")
            If closeBracketPos > 0 Then
                Dim sectionHeader As String
                sectionHeader = Left(line, closeBracketPos)
                If UCase(sectionHeader) = UCase(targetSection) Then
                    inSection = True
                Else
                    If inSection Then
                        ' 次のセクションに入ったので終了
                        Exit For
                    End If
                    inSection = False
                End If
            End If
        ElseIf inSection And line <> "" And Left(line, 1) <> ";" And InStr(line, "=") > 0 Then
            ' キー=値の行
            Dim semicolonPos As Long
            semicolonPos = InStr(line, ";")
            
            Dim keyValuePart As String
            If semicolonPos > 0 Then
                keyValuePart = Trim(Left(line, semicolonPos - 1))
            Else
                keyValuePart = line
            End If
            
            Dim parts() As String
            parts = Split(keyValuePart, "=", 2)
            
            If UBound(parts) >= 1 And UCase(Trim(parts(0))) = UCase(keyName) Then
                lineToRemove = i
                Exit For
            End If
        End If
    Next i
    
    ' キーを削除
    If lineToRemove > 0 Then
        m_lines.Remove lineToRemove
        RemoveKey = True
    Else
        RemoveKey = False
    End If
    
    Exit Function
    
ErrorHandler:
    RemoveKey = False
End Function

' 全セクション名を取得
Public Function GetSectionNames() As String()
    On Error GoTo ErrorHandler
    
    If Not m_isLoaded Then
        GetSectionNames = Split("", ",")
        Exit Function
    End If
    
    Dim sectionList As Collection
    Set sectionList = New Collection
    
    Dim i As Long
    For i = 1 To m_lines.Count
        Dim line As String
        line = Trim(m_lines.Item(i))
        
        If Left(line, 1) = "[" Then
            Dim closeBracketPos As Long
            closeBracketPos = InStr(line, "]")
            If closeBracketPos > 0 Then
                Dim sectionName As String
                sectionName = Mid(line, 2, closeBracketPos - 2)  ' [と]を除去
                
                ' 重複チェック（大文字・小文字を区別しない）
                Dim alreadyExists As Boolean
                alreadyExists = False
                Dim j As Long
                For j = 1 To sectionList.Count
                    If UCase(sectionList.Item(j)) = UCase(sectionName) Then
                        alreadyExists = True
                        Exit For
                    End If
                Next j
                
                If Not alreadyExists Then
                    sectionList.Add sectionName
                End If
            End If
        End If
    Next i
    
    Dim sections() As String
    ReDim sections(0 To sectionList.Count - 1)
    
    For i = 1 To sectionList.Count
        sections(i - 1) = sectionList.Item(i)
    Next i
    
    GetSectionNames = sections
    Exit Function
    
ErrorHandler:
    GetSectionNames = Split("", ",")
End Function

'==============================================================================
' 旧バージョンとの互換性維持のための関数
'==============================================================================
Public Function GetValue(sectionName As String, keyName As String, Optional defaultValue As String = "") As String
    GetValue = ReadValue(sectionName, keyName, defaultValue)
End Function

Public Function GetInteger(sectionName As String, keyName As String, Optional defaultValue As Long = 0) As Long
    GetInteger = ReadValueAsInteger(sectionName, keyName, defaultValue)
End Function

Public Function GetDouble(sectionName As String, keyName As String, Optional defaultValue As Double = 0#) As Double
    GetDouble = ReadValueAsDouble(sectionName, keyName, defaultValue)
End Function

Public Function GetBoolean(sectionName As String, keyName As String, Optional defaultValue As Boolean = False) As Boolean
    GetBoolean = ReadValueAsBoolean(sectionName, keyName, defaultValue)
End Function

Public Function SetValue(sectionName As String, keyName As String, value As String) As Boolean
    SetValue = WriteValue(sectionName, keyName, value)
End Function

Public Function SetInteger(sectionName As String, keyName As String, value As Long) As Boolean
    SetInteger = WriteValueAsInteger(sectionName, keyName, value)
End Function

Public Function SetDouble(sectionName As String, keyName As String, value As Double) As Boolean
    SetDouble = WriteValueAsDouble(sectionName, keyName, value)
End Function

'==============================================================================
' プライベートメソッド
'==============================================================================

' ファイルから読み込み
Private Function LoadFromFile() As Boolean
    On Error GoTo ErrorHandler
    
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    
    stream.Type = 2  ' テキストストリーム
    stream.Charset = m_encoding
    stream.Open
    stream.LoadFromFile m_filePath
    
    Dim content As String
    content = stream.ReadText
    stream.Close
    
    ParseContent content
    LoadFromFile = True
    
    Exit Function
    
ErrorHandler:
    LoadFromFile = False
    If Not stream Is Nothing Then
        If stream.State = 1 Then stream.Close
    End If
End Function

' ファイルに保存
Private Function SaveToFile() As Boolean
    On Error GoTo ErrorHandler
    
    Dim content As String
    content = BuildContent()
    
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    
    stream.Type = 2  ' テキストストリーム
    stream.Charset = m_encoding
    stream.Open
    stream.WriteText content
    stream.SaveToFile m_filePath, 2  ' 上書き保存
    stream.Close
    
    SaveToFile = True
    
    Exit Function
    
ErrorHandler:
    SaveToFile = False
    If Not stream Is Nothing Then
        If stream.State = 1 Then stream.Close
    End If
End Function

' コンテンツを解析
Private Sub ParseContent(content As String)
    Set m_lines = New Collection
    
    Dim lines() As String
    lines = Split(Replace(content, vbCrLf, vbLf), vbLf)
    
    ' 全ての行をそのまま保持
    Dim i As Long
    For i = 0 To UBound(lines)
        m_lines.Add lines(i)
    Next i
End Sub

' コンテンツを構築
Private Function BuildContent() As String
    Dim result As String
    
    Dim i As Long
    For i = 1 To m_lines.Count
        result = result & m_lines.Item(i)
        
        ' 最終行以外は常に改行を追加
        If i < m_lines.Count Then
            result = result & vbCrLf
        End If
    Next i
    
    ' 最終行にCR/LFがない場合のみ追加（内容が空でない場合）
    If m_lines.Count > 0 And result <> "" And Not (Right(result, 2) = vbCrLf Or Right(result, 1) = vbLf) Then
        result = result & vbCrLf
    End If
    
    BuildContent = result
End Function



'==============================================================================
' クラス終了処理
'==============================================================================

Private Sub Class_Terminate()
    Set m_lines = Nothing
End Sub

