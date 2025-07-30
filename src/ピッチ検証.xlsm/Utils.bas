Attribute VB_Name = "Utils"
Option Explicit

'-----------------------------------------------------
'   GCD(Greatest Common Divisor) 最大公約数
'
'   ユークリッドの互除法で求める(再帰版)
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
'   LCM(Least Common Multiple) 最小公倍数
'-----------------------------------------------------
Public Function lcm(ByVal m As Long, ByVal N As Long) As Long

    If m = 0 Or N = 0 Then
        lcm = 0
    Else
        lcm = (m / gcd(m, N)) * N
    End If

End Function

'浮動小数を単純連分数に
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
    
    '最後が1の場合は1/1になるのでその前の値を+1して次数を減らす
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

'連分数を分数に
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
' ファイル選択ダイアログ
'
' 引数
'   Title       : ダイアログタイトル
'   initFolder  : 初期表示フォルダー
'   fileType    : ファイル種類(例:Excelファイル)
'   filter      : 拡張子フィルター(例:*.xlsx)
'----------------------------------------------------------------------------------------------------
Function OpenFileDialog(ByVal Title As String, ByVal initFolder As String, ByVal fileType As String, ByVal filter As String) As String
    Dim fd As FileDialog
    Dim filePath As String

    ' FileDialog オブジェクト取得
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = Title
        .InitialFileName = initFolder
        .filters.Clear
        .filters.Add fileType, filter
        .AllowMultiSelect = False

        ' ダイアログ表示
        If .Show = -1 Then
            ' 選択されたファイルを取得
            OpenFileDialog = .SelectedItems(1)
        Else
            OpenFileDialog = ""
        End If
    End With

    Set fd = Nothing

End Function

'----------------------------------------------------------------------------------------------------
' ファイル選択ダイアログ
'
' 引数
'   Title               : ダイアログタイトル
'   AllowMultiSelect    : True=複数ファイル選択可、False=単一ファイル
'   filters             : 拡張子フィルターの配列
'----------------------------------------------------------------------------------------------------
Function ShowFileDialogWithFilters( _
    Optional Title As String = "ファイルを選択してください", _
    Optional AllowMultiSelect As Boolean = False, _
    Optional filters As Variant = Empty _
) As Variant
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = Title
        .AllowMultiSelect = AllowMultiSelect
        .filters.Clear
        
        ' フィルターが指定されていれば追加
        If Not IsEmpty(filters) Then
            Dim i As Long
            For i = LBound(filters) To UBound(filters)
                ' 各フィルターは配列形式: Array("表示名", "*.ext1;*.ext2")
                If IsArray(filters(i)) And UBound(filters(i)) = 1 Then
                    .filters.Add filters(i)(0), filters(i)(1)
                End If
            Next i
        End If
        
        ' ダイアログ表示
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


'指定したシートの指定した行から最後までクリアする
Sub ClearFromRow(ws As Worksheet, startRow As Long)
    Dim lastRow As Long
    
    ' 最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    
    ' 指定した行から最終行までをクリア
    If lastRow >= startRow Then
        ws.Range(ws.Cells(startRow, 1), ws.Cells(lastRow, ws.Columns.Count)).ClearContents
    End If
End Sub

