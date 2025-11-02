Option Explicit

' ============================================
' 車割自動作成マクロ v1.2 (互換性改善版)
' ============================================
' 使い方:
' 1. 「メンバー情報」シートにデータを入力
' 2. このマクロを実行（Alt+F8 → GenerateKurumawari）
' 3. 「車割結果」シートに結果が出力されます
'
' 更新履歴:
' v1.2 (2025/11/02) - ArrayListを使わない互換性改善版
' v1.1 (2025/11/02) - 統計情報の追加、エクスポート機能の追加
' v1.0 (2025/11/02) - 初版リリース
' ============================================

' 定数定義
Const MAX_CAPACITY As Integer = 5 ' 1台あたりの最大乗車人数（運転手含む）
Const MAX_PASSENGERS As Integer = 4 ' 同乗者の最大人数

' メンバー情報を格納する構造体
Type MemberInfo
    Name As String
    CanDrive As Boolean
    OutboundDate As String
    OutboundTime As String
    OutboundLocation As String
    ReturnDate As String
    ReturnTime As String
    ReturnLocation As String
End Type

' 車割グループ
Type CarGroup
    GroupDate As String
    GroupTime As String
    GroupLocation As String
    Direction As String ' "行き" or "帰り"
    Members() As String
    MemberCount As Integer
    Drivers() As String
    DriverCount As Integer
End Type

' ============================================
' メイン処理
' ============================================
Sub GenerateKurumawari()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' シートの存在確認
    Dim wsData As Worksheet, wsResult As Worksheet
    
    On Error Resume Next
    Set wsData = ThisWorkbook.Worksheets("メンバー情報")
    Set wsResult = ThisWorkbook.Worksheets("車割結果")
    On Error GoTo ErrorHandler
    
    If wsData Is Nothing Then
        MsgBox "「メンバー情報」シートが見つかりません。", vbExclamation
        Exit Sub
    End If
    
    If wsResult Is Nothing Then
        MsgBox "「車割結果」シートが見つかりません。", vbExclamation
        Exit Sub
    End If
    
    ' データ読み込み
    Dim members() As MemberInfo
    Dim memberCount As Long
    If Not LoadMemberData(wsData, members, memberCount) Then
        MsgBox "メンバー情報の読み込みに失敗しました。", vbExclamation
        Exit Sub
    End If
    
    ' 車割グループの作成
    Dim groups() As CarGroup
    Dim groupCount As Long
    CreateCarGroups members, memberCount, groups, groupCount
    
    ' 結果を出力
    OutputResults wsResult, groups, groupCount
    
    ' 統計情報を追加
    AddStatistics wsResult, groups, groupCount, members, memberCount
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    ' 完了メッセージに統計情報を含める
    MsgBox "車割の作成が完了しました！" & vbCrLf & vbCrLf & _
           "総台数: " & groupCount & " 台" & vbCrLf & _
           "総人数: " & memberCount & " 人", _
           vbInformation, "車割作成完了"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "エラーが発生しました: " & Err.Description & vbCrLf & _
           "エラー番号: " & Err.Number, vbCritical
End Sub

' ============================================
' メンバー情報の読み込み
' ============================================
Function LoadMemberData(ws As Worksheet, ByRef members() As MemberInfo, ByRef count As Long) As Boolean
    On Error GoTo ErrorHandler
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If lastRow < 2 Then
        LoadMemberData = False
        Exit Function
    End If
    
    Dim i As Long
    count = 0
    
    ' メンバー数をカウント
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value <> "" Then
            count = count + 1
        End If
    Next i
    
    If count = 0 Then
        LoadMemberData = False
        Exit Function
    End If
    
    ' 配列のサイズを設定
    ReDim members(1 To count)
    
    ' データを読み込み
    Dim idx As Long
    idx = 1
    
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value <> "" Then
            With members(idx)
                .Name = Trim(CStr(ws.Cells(i, 1).Value))
                .OutboundDate = Trim(CStr(ws.Cells(i, 2).Value))
                .OutboundTime = Trim(CStr(ws.Cells(i, 3).Value))
                .OutboundLocation = Trim(CStr(ws.Cells(i, 4).Value))
                .ReturnDate = Trim(CStr(ws.Cells(i, 5).Value))
                .ReturnTime = Trim(CStr(ws.Cells(i, 6).Value))
                .ReturnLocation = Trim(CStr(ws.Cells(i, 7).Value))
                .CanDrive = (Trim(CStr(ws.Cells(i, 8).Value)) = "○")
            End With
            idx = idx + 1
        End If
    Next i
    
    LoadMemberData = True
    Exit Function
    
ErrorHandler:
    LoadMemberData = False
End Function

' ============================================
' 車割グループの作成（シンプル版）
' ============================================
Sub CreateCarGroups(members() As MemberInfo, memberCount As Long, _
                     ByRef groups() As CarGroup, ByRef groupCount As Long)
    
    ' 一時的なグループ情報を格納
    Dim tempGroups() As CarGroup
    ReDim tempGroups(1 To memberCount * 2) ' 最大でメンバー数×2（行き帰り）
    
    Dim tempGroupCount As Long
    tempGroupCount = 0
    
    Dim i As Long
    Dim foundIdx As Long
    Dim key As String
    
    ' 行きと帰りでグループ化
    For i = 1 To memberCount
        ' 行きのグループ
        If members(i).OutboundDate <> "" Then
            key = members(i).OutboundDate & "|" & members(i).OutboundTime & "|" & _
                  members(i).OutboundLocation
            
            foundIdx = FindGroupByKey(tempGroups, tempGroupCount, key, "行き")
            
            If foundIdx = 0 Then
                ' 新しいグループを作成
                tempGroupCount = tempGroupCount + 1
                With tempGroups(tempGroupCount)
                    .GroupDate = members(i).OutboundDate
                    .GroupTime = members(i).OutboundTime
                    .GroupLocation = members(i).OutboundLocation
                    .Direction = "行き"
                    .MemberCount = 0
                    .DriverCount = 0
                    ReDim .Members(1 To memberCount)
                    ReDim .Drivers(1 To memberCount)
                End With
                foundIdx = tempGroupCount
            End If
            
            ' メンバーを追加
            With tempGroups(foundIdx)
                .MemberCount = .MemberCount + 1
                .Members(.MemberCount) = members(i).Name
                If members(i).CanDrive Then
                    .DriverCount = .DriverCount + 1
                    .Drivers(.DriverCount) = members(i).Name
                End If
            End With
        End If
        
        ' 帰りのグループ
        If members(i).ReturnDate <> "" Then
            key = members(i).ReturnDate & "|" & members(i).ReturnTime & "|" & _
                  members(i).ReturnLocation
            
            foundIdx = FindGroupByKey(tempGroups, tempGroupCount, key, "帰り")
            
            If foundIdx = 0 Then
                ' 新しいグループを作成
                tempGroupCount = tempGroupCount + 1
                With tempGroups(tempGroupCount)
                    .GroupDate = members(i).ReturnDate
                    .GroupTime = members(i).ReturnTime
                    .GroupLocation = members(i).ReturnLocation
                    .Direction = "帰り"
                    .MemberCount = 0
                    .DriverCount = 0
                    ReDim .Members(1 To memberCount)
                    ReDim .Drivers(1 To memberCount)
                End With
                foundIdx = tempGroupCount
            End If
            
            ' メンバーを追加
            With tempGroups(foundIdx)
                .MemberCount = .MemberCount + 1
                .Members(.MemberCount) = members(i).Name
                If members(i).CanDrive Then
                    .DriverCount = .DriverCount + 1
                    .Drivers(.DriverCount) = members(i).Name
                End If
            End With
        End If
    Next i
    
    ' 車の台数を計算して最終的なグループを作成
    groupCount = 0
    For i = 1 To tempGroupCount
        groupCount = groupCount + GetRequiredCars(tempGroups(i).MemberCount)
    Next i
    
    ReDim groups(1 To groupCount)
    
    Dim groupIdx As Long
    groupIdx = 1
    
    ' 各グループを車に割り当て
    For i = 1 To tempGroupCount
        Dim numCars As Integer
        numCars = GetRequiredCars(tempGroups(i).MemberCount)
        
        Dim carIdx As Integer
        For carIdx = 1 To numCars
            groups(groupIdx).GroupDate = tempGroups(i).GroupDate
            groups(groupIdx).GroupTime = tempGroups(i).GroupTime
            groups(groupIdx).GroupLocation = tempGroups(i).GroupLocation
            groups(groupIdx).Direction = tempGroups(i).Direction
            
            AssignMembersToCar groups(groupIdx), tempGroups(i), carIdx, numCars
            
            groupIdx = groupIdx + 1
        Next carIdx
    Next i
    
End Sub

' ============================================
' グループを検索
' ============================================
Function FindGroupByKey(groups() As CarGroup, count As Long, key As String, direction As String) As Long
    Dim i As Long
    Dim checkKey As String
    
    For i = 1 To count
        checkKey = groups(i).GroupDate & "|" & groups(i).GroupTime & "|" & groups(i).GroupLocation
        If checkKey = key And groups(i).Direction = direction Then
            FindGroupByKey = i
            Exit Function
        End If
    Next i
    
    FindGroupByKey = 0
End Function

' ============================================
' 必要な車の台数を計算
' ============================================
Function GetRequiredCars(memberCount As Integer) As Integer
    GetRequiredCars = WorksheetFunction.RoundUp(memberCount / MAX_CAPACITY, 0)
End Function

' ============================================
' 車にメンバーを割り当て
' ============================================
Sub AssignMembersToCar(ByRef group As CarGroup, ByRef sourceGroup As CarGroup, _
                        carNum As Integer, totalCars As Integer)
    
    Dim memberCount As Integer
    memberCount = sourceGroup.MemberCount
    
    ' 各車に割り当てる人数を計算
    Dim basePerCar As Integer
    Dim extraMembers As Integer
    basePerCar = memberCount \ totalCars
    extraMembers = memberCount Mod totalCars
    
    Dim membersInThisCar As Integer
    If carNum <= extraMembers Then
        membersInThisCar = basePerCar + 1
    Else
        membersInThisCar = basePerCar
    End If
    
    ' 開始インデックスを計算
    Dim startIdx As Integer
    If carNum <= extraMembers Then
        startIdx = (carNum - 1) * (basePerCar + 1)
    Else
        startIdx = extraMembers * (basePerCar + 1) + (carNum - extraMembers - 1) * basePerCar
    End If
    
    ' メンバーを配列に格納
    ReDim group.Members(1 To membersInThisCar)
    group.MemberCount = membersInThisCar
    
    Dim i As Integer
    For i = 1 To membersInThisCar
        group.Members(i) = sourceGroup.Members(startIdx + i)
    Next i
    
    ' 運転手を決定（このグループ内に運転可能な人がいる場合）
    Dim driverFound As Boolean
    driverFound = False
    
    For i = 1 To membersInThisCar
        If IsInDriverList(group.Members(i), sourceGroup.Drivers, sourceGroup.DriverCount) Then
            ReDim group.Drivers(1 To 1)
            group.Drivers(1) = group.Members(i)
            group.DriverCount = 1
            driverFound = True
            Exit For
        End If
    Next i
    
    ' 運転手が見つからない場合は最初の人を運転手に
    If Not driverFound Then
        ReDim group.Drivers(1 To 1)
        group.Drivers(1) = group.Members(1) & " (要確認)"
        group.DriverCount = 1
    End If
    
End Sub

' ============================================
' 運転手リストに存在するか確認
' ============================================
Function IsInDriverList(memberName As String, drivers() As String, driverCount As Integer) As Boolean
    Dim i As Integer
    IsInDriverList = False
    
    For i = 1 To driverCount
        If drivers(i) = memberName Then
            IsInDriverList = True
            Exit Function
        End If
    Next i
End Function

' ============================================
' 結果をシートに出力
' ============================================
Sub OutputResults(ws As Worksheet, groups() As CarGroup, count As Long)
    ' シートをクリア
    ws.Cells.Clear
    
    ' ヘッダー行を作成
    ws.Cells(1, 1).Value = "日"
    ws.Cells(1, 2).Value = "時"
    ws.Cells(1, 3).Value = "場所"
    ws.Cells(1, 4).Value = "運転手"
    
    Dim col As Integer
    For col = 1 To MAX_PASSENGERS
        ws.Cells(1, 4 + col).Value = "同乗者" & col
    Next col
    
    ' ヘッダーの書式設定
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, 4 + MAX_PASSENGERS))
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
        .HorizontalAlignment = xlCenter
    End With
    
    ' データを出力
    Dim row As Long
    row = 2
    
    Dim i As Long
    For i = 1 To count
        ws.Cells(row, 1).Value = groups(i).GroupDate
        ws.Cells(row, 2).Value = groups(i).GroupTime
        ws.Cells(row, 3).Value = groups(i).GroupLocation
        
        ' 運転手
        If groups(i).DriverCount >= 1 Then
            ws.Cells(row, 4).Value = groups(i).Drivers(1)
        End If
        
        ' 同乗者
        Dim passengerCol As Integer
        passengerCol = 1
        
        Dim j As Integer
        For j = 1 To groups(i).MemberCount
            ' 運転手以外を同乗者として出力
            If groups(i).Members(j) <> ws.Cells(row, 4).Value Then
                If passengerCol <= MAX_PASSENGERS Then
                    ws.Cells(row, 4 + passengerCol).Value = groups(i).Members(j)
                    passengerCol = passengerCol + 1
                End If
            End If
        Next j
        
        row = row + 1
    Next i
    
    ' 列幅を自動調整
    ws.Columns("A:H").AutoFit
    
    ' 罫線を追加
    With ws.Range(ws.Cells(1, 1), ws.Cells(row - 1, 4 + MAX_PASSENGERS))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    
    ' シートをアクティブに
    ws.Activate
    ws.Range("A1").Select
    
End Sub

' ============================================
' 統計情報の追加
' ============================================
Sub AddStatistics(ws As Worksheet, groups() As CarGroup, groupCount As Long, _
                   members() As MemberInfo, memberCount As Long)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' 統計情報を追加（3行空けて）
    Dim statsRow As Long
    statsRow = lastRow + 3
    
    ' タイトル
    ws.Cells(statsRow, 1).Value = "【統計情報】"
    ws.Cells(statsRow, 1).Font.Bold = True
    ws.Cells(statsRow, 1).Font.Size = 12
    
    statsRow = statsRow + 1
    
    ' 総台数
    ws.Cells(statsRow, 1).Value = "総台数:"
    ws.Cells(statsRow, 2).Value = groupCount & " 台"
    statsRow = statsRow + 1
    
    ' 総人数
    ws.Cells(statsRow, 1).Value = "総人数:"
    ws.Cells(statsRow, 2).Value = memberCount & " 人"
    statsRow = statsRow + 1
    
    ' 運転可能な人数
    Dim driverCount As Integer
    driverCount = 0
    Dim i As Long
    For i = 1 To memberCount
        If members(i).CanDrive Then
            driverCount = driverCount + 1
        End If
    Next i
    ws.Cells(statsRow, 1).Value = "運転可能:"
    ws.Cells(statsRow, 2).Value = driverCount & " 人"
    statsRow = statsRow + 1
    
    ' 1台あたりの平均乗車人数
    Dim avgPerCar As Double
    avgPerCar = CDbl(memberCount) / CDbl(groupCount)
    ws.Cells(statsRow, 1).Value = "平均乗車人数:"
    ws.Cells(statsRow, 2).Value = Format(avgPerCar, "0.0") & " 人/台"
    
    ' 統計情報の書式設定
    With ws.Range(ws.Cells(lastRow + 3, 1), ws.Cells(statsRow, 2))
        .Font.Color = RGB(0, 0, 139)
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlMedium
    End With
    
End Sub

' ============================================
' CSV形式でエクスポート（オプション機能）
' ============================================
Sub ExportToCSV()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("車割結果")
    
    ' ファイル保存ダイアログ
    Dim filePath As String
    filePath = Application.GetSaveAsFilename( _
        InitialFileName:="車割結果_" & Format(Date, "yyyymmdd") & ".csv", _
        FileFilter:="CSV Files (*.csv), *.csv", _
        Title:="CSVファイルとして保存")
    
    If filePath = "False" Then Exit Sub
    
    ' CSVとして保存
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' 統計情報の行を除外（【統計情報】が見つかるまで）
    Dim exportRow As Long
    exportRow = lastRow
    Dim i As Long
    For i = 1 To lastRow
        If InStr(ws.Cells(i, 1).Value, "【統計情報】") > 0 Then
            exportRow = i - 2
            Exit For
        End If
    Next i
    
    ' エクスポート範囲をコピーして新しいブックに貼り付け
    Dim newWb As Workbook
    Set newWb = Workbooks.Add
    
    ws.Range(ws.Cells(1, 1), ws.Cells(exportRow, 8)).Copy
    newWb.Sheets(1).Range("A1").PasteSpecial xlPasteValues
    
    ' CSVとして保存
    Application.DisplayAlerts = False
    newWb.SaveAs Filename:=filePath, FileFormat:=xlCSV
    newWb.Close False
    Application.DisplayAlerts = True
    
    MsgBox "CSVファイルを保存しました: " & vbCrLf & filePath, vbInformation
    
    Exit Sub
    
ErrorHandler:
    MsgBox "エクスポートに失敗しました: " & Err.Description, vbCritical
End Sub
