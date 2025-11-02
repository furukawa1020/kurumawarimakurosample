Attribute VB_Name = "KurumawariMacro"
Option Explicit

' ============================================
' 車割自動作成マクロ v1.0
' ============================================
' 使い方:
' 1. 「メンバー情報」シートにデータを入力
' 2. このマクロを実行（Alt+F8 → GenerateKurumawari）
' 3. 「車割結果」シートに結果が出力されます
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
    Drivers() As String
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
    If Not LoadMemberData(wsData, members) Then
        MsgBox "メンバー情報の読み込みに失敗しました。", vbExclamation
        Exit Sub
    End If
    
    ' 車割グループの作成
    Dim groups() As CarGroup
    CreateCarGroups members, groups
    
    ' 結果を出力
    OutputResults wsResult, groups
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "車割の作成が完了しました！", vbInformation
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
End Sub

' ============================================
' メンバー情報の読み込み
' ============================================
Function LoadMemberData(ws As Worksheet, ByRef members() As MemberInfo) As Boolean
    On Error GoTo ErrorHandler
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If lastRow < 2 Then
        LoadMemberData = False
        Exit Function
    End If
    
    Dim i As Long, count As Long
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
                .Name = Trim(ws.Cells(i, 1).Value)
                .OutboundDate = Trim(ws.Cells(i, 2).Value)
                .OutboundTime = Trim(ws.Cells(i, 3).Value)
                .OutboundLocation = Trim(ws.Cells(i, 4).Value)
                .ReturnDate = Trim(ws.Cells(i, 5).Value)
                .ReturnTime = Trim(ws.Cells(i, 6).Value)
                .ReturnLocation = Trim(ws.Cells(i, 7).Value)
                .CanDrive = (Trim(ws.Cells(i, 8).Value) = "○")
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
' 車割グループの作成
' ============================================
Sub CreateCarGroups(members() As MemberInfo, ByRef groups() As CarGroup)
    Dim groupList As Object
    Set groupList = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    Dim key As String
    
    ' 行きと帰りでグループ化
    For i = LBound(members) To UBound(members)
        ' 行きのグループ
        If members(i).OutboundDate <> "" Then
            key = members(i).OutboundDate & "|" & members(i).OutboundTime & "|" & _
                  members(i).OutboundLocation & "|行き"
            
            If Not groupList.Exists(key) Then
                groupList.Add key, CreateObject("Scripting.Dictionary")
                groupList(key)("Date") = members(i).OutboundDate
                groupList(key)("Time") = members(i).OutboundTime
                groupList(key)("Location") = members(i).OutboundLocation
                groupList(key)("Direction") = "行き"
                groupList(key)("Members") = CreateObject("System.Collections.ArrayList")
                groupList(key)("Drivers") = CreateObject("System.Collections.ArrayList")
            End If
            
            groupList(key)("Members").Add members(i).Name
            If members(i).CanDrive Then
                groupList(key)("Drivers").Add members(i).Name
            End If
        End If
        
        ' 帰りのグループ
        If members(i).ReturnDate <> "" Then
            key = members(i).ReturnDate & "|" & members(i).ReturnTime & "|" & _
                  members(i).ReturnLocation & "|帰り"
            
            If Not groupList.Exists(key) Then
                groupList.Add key, CreateObject("Scripting.Dictionary")
                groupList(key)("Date") = members(i).ReturnDate
                groupList(key)("Time") = members(i).ReturnTime
                groupList(key)("Location") = members(i).ReturnLocation
                groupList(key)("Direction") = "帰り"
                groupList(key)("Members") = CreateObject("System.Collections.ArrayList")
                groupList(key)("Drivers") = CreateObject("System.Collections.ArrayList")
            End If
            
            groupList(key)("Members").Add members(i).Name
            If members(i).CanDrive Then
                groupList(key)("Drivers").Add members(i).Name
            End If
        End If
    Next i
    
    ' グループを配列に変換
    Dim groupCount As Long
    groupCount = 0
    
    Dim groupKey As Variant
    For Each groupKey In groupList.Keys
        groupCount = groupCount + GetRequiredCars(groupList(groupKey)("Members").Count)
    Next groupKey
    
    ReDim groups(1 To groupCount)
    
    Dim groupIdx As Long
    groupIdx = 1
    
    ' 各グループを車に割り当て
    For Each groupKey In groupList.Keys
        Dim grp As Object
        Set grp = groupList(groupKey)
        
        Dim memberArray() As String
        Dim driverArray() As String
        
        memberArray = CollectionToArray(grp("Members"))
        driverArray = CollectionToArray(grp("Drivers"))
        
        ' 車の台数を計算
        Dim numCars As Integer
        numCars = GetRequiredCars(UBound(memberArray) - LBound(memberArray) + 1)
        
        ' 各車にメンバーを割り当て
        Dim carIdx As Integer
        For carIdx = 1 To numCars
            groups(groupIdx).GroupDate = grp("Date")
            groups(groupIdx).GroupTime = grp("Time")
            groups(groupIdx).GroupLocation = grp("Location")
            groups(groupIdx).Direction = grp("Direction")
            
            AssignMembersToCAR groups(groupIdx), memberArray, driverArray, carIdx, numCars
            
            groupIdx = groupIdx + 1
        Next carIdx
    Next groupKey
    
End Sub

' ============================================
' 必要な車の台数を計算
' ============================================
Function GetRequiredCars(memberCount As Integer) As Integer
    GetRequiredCars = WorksheetFunction.RoundUp(memberCount / MAX_CAPACITY, 0)
End Function

' ============================================
' 車にメンバーを割り当て
' ============================================
Sub AssignMembersToCAR(ByRef group As CarGroup, ByRef memberArray() As String, _
                        ByRef driverArray() As String, carNum As Integer, totalCars As Integer)
    
    Dim driverCount As Integer
    driverCount = UBound(driverArray) - LBound(driverArray) + 1
    
    Dim memberCount As Integer
    memberCount = UBound(memberArray) - LBound(memberArray) + 1
    
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
    
    Dim i As Integer
    For i = 1 To membersInThisCar
        group.Members(i) = memberArray(LBound(memberArray) + startIdx + i - 1)
    Next i
    
    ' 運転手を決定（このグループ内に運転可能な人がいる場合）
    Dim driverFound As Boolean
    driverFound = False
    
    For i = 1 To membersInThisCar
        If IsInArray(group.Members(i), driverArray) Then
            ReDim group.Drivers(1 To 1)
            group.Drivers(1) = group.Members(i)
            driverFound = True
            Exit For
        End If
    Next i
    
    ' 運転手が見つからない場合は最初の人を運転手に
    If Not driverFound Then
        ReDim group.Drivers(1 To 1)
        group.Drivers(1) = group.Members(1) & " (要確認)"
    End If
    
End Sub

' ============================================
' 配列内検索
' ============================================
Function IsInArray(val As String, arr() As String) As Boolean
    Dim i As Long
    IsInArray = False
    
    On Error Resume Next
    For i = LBound(arr) To UBound(arr)
        If arr(i) = val Then
            IsInArray = True
            Exit Function
        End If
    Next i
End Function

' ============================================
' コレクションを配列に変換
' ============================================
Function CollectionToArray(col As Object) As String()
    Dim arr() As String
    Dim i As Long
    
    If col.Count = 0 Then
        ReDim arr(1 To 1)
        arr(1) = ""
        CollectionToArray = arr
        Exit Function
    End If
    
    ReDim arr(1 To col.Count)
    
    For i = 1 To col.Count
        arr(i) = col(i - 1)
    Next i
    
    CollectionToArray = arr
End Function

' ============================================
' 結果をシートに出力
' ============================================
Sub OutputResults(ws As Worksheet, groups() As CarGroup)
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
    For i = LBound(groups) To UBound(groups)
        ws.Cells(row, 1).Value = groups(i).GroupDate
        ws.Cells(row, 2).Value = groups(i).GroupTime
        ws.Cells(row, 3).Value = groups(i).GroupLocation
        
        ' 運転手
        If UBound(groups(i).Drivers) >= 1 Then
            ws.Cells(row, 4).Value = groups(i).Drivers(1)
        End If
        
        ' 同乗者
        Dim passengerCol As Integer
        passengerCol = 1
        
        Dim j As Long
        For j = LBound(groups(i).Members) To UBound(groups(i).Members)
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
