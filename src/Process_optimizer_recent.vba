Option Explicit

Sub Search()
    Dim wsSearch As Worksheet, ws As Worksheet
    Dim searchWord As String
    Dim lastRow As Long, i As Long, j As Long, resRow As Long
    Dim foundCount As Integer
    Dim rules As Variant
    
    ' 시트 설정 및 검색어 확인
    Set wsSearch = ThisWorkbook.Sheets("검색")
    searchWord = wsSearch.Range("B2").Value
    
    If searchWord = "" Then
        MsgBox "검색어를 입력해주세요."
        Exit Sub
    End If

    ' 성능 최적화 설정
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' 이전 결과, 기존 서식 및 총 전력소비량(G2) 초기화
    wsSearch.Rows("5:" & wsSearch.Rows.Count).Delete
    wsSearch.Range("G2").Value = ""
    
    resRow = 5

    ' [생산 품목] 검색
    wsSearch.Cells(resRow, 1).Value = "▶ [" & searchWord & "]을(를) 생산하는 공정"
    wsSearch.Cells(resRow, 1).Font.Bold = True
    resRow = resRow + 1
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> wsSearch.Name Then
            rules = GetSheetRules(ws)
            
            If Not IsEmpty(rules) Then
                lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
                foundCount = 0
                
                Dim colOut1 As String: colOut1 = CStr(rules(4))
                Dim colOut2 As String: colOut2 = CStr(rules(5))
                
                For i = 2 To lastRow
                    Dim isOutMatch As Boolean: isOutMatch = False
                    
                    If InStr(ws.Cells(i, colOut1).MergeArea(1).Value, searchWord) > 0 Then
                        isOutMatch = True
                    ElseIf colOut2 <> "" Then
                        If InStr(ws.Cells(i, colOut2).MergeArea(1).Value, searchWord) > 0 Then isOutMatch = True
                    End If
                    
                    If isOutMatch Then
                        If foundCount = 0 Then
                            ws.Rows(1).Copy Destination:=wsSearch.Rows(resRow)
                            wsSearch.Cells(resRow, 2).Value = "생산 시설명(" & ws.Name & ")"
                            resRow = resRow + 1
                        End If
                        
                        ws.Rows(i).Copy Destination:=wsSearch.Rows(resRow)
                        For j = 1 To 15
                            wsSearch.Cells(resRow, j).Value = ws.Cells(i, j).MergeArea(1).Value
                        Next j
                        
                        resRow = resRow + 1
                        foundCount = foundCount + 1
                    End If
                Next i
            End If
        End If
    Next ws

    resRow = resRow + 2

    ' [소모 재료] 검색
    wsSearch.Cells(resRow, 1).Value = "▶ [" & searchWord & "]을(를) 재료로 소모하는 공정"
    wsSearch.Cells(resRow, 1).Font.Bold = True
    resRow = resRow + 1
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> wsSearch.Name Then
            rules = GetSheetRules(ws)
            
            If Not IsEmpty(rules) Then
                lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
                foundCount = 0
                
                Dim colIn1 As String: colIn1 = CStr(rules(0))
                Dim colIn2 As String: colIn2 = CStr(rules(1))
                
                For i = 2 To lastRow
                    Dim isInMatch As Boolean: isInMatch = False
                    
                    If InStr(ws.Cells(i, colIn1).MergeArea(1).Value, searchWord) > 0 Then
                        isInMatch = True
                    ElseIf colIn2 <> "" Then
                        If InStr(ws.Cells(i, colIn2).MergeArea(1).Value, searchWord) > 0 Then isInMatch = True
                    End If
                    
                    If isInMatch Then
                        If foundCount = 0 Then
                            ws.Rows(1).Copy Destination:=wsSearch.Rows(resRow)
                            wsSearch.Cells(resRow, 2).Value = "생산 시설명(" & ws.Name & ")"
                            resRow = resRow + 1
                        End If
                        
                        ws.Rows(i).Copy Destination:=wsSearch.Rows(resRow)
                        For j = 1 To 15
                            wsSearch.Cells(resRow, j).Value = ws.Cells(i, j).MergeArea(1).Value
                        Next j
                        
                        resRow = resRow + 1
                        foundCount = foundCount + 1
                    End If
                Next i
            End If
        End If
    Next ws

    ' 최적화 설정 복구
    Application.CutCopyMode = False
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    wsSearch.Columns.AutoFit
    MsgBox "검색 완료"
End Sub

Sub SearchFullRecipe()
    Dim wsSearch As Worksheet: Set wsSearch = ThisWorkbook.Sheets("검색")
    Dim itemName As String
    Dim ws As Worksheet
    Dim i As Long
    
    itemName = Trim(Selection.Cells(1, 1).MergeArea(1).Text)
    wsSearch.Range("E2").Value = itemName
    
    If itemName = "" Then
        MsgBox "분석할 아이템 명이 적힌 셀을 선택해주세요."
        Exit Sub
    End If

    Application.ScreenUpdating = False
    InitializeSearchSheet wsSearch

    Dim targetPPS As Double
    Dim extraQ As Object: Set extraQ = CreateObject("System.Collections.ArrayList")
    Dim checkedList As Object: Set checkedList = CreateObject("Scripting.Dictionary")
    
    ' 분석 대상 시트 수집
    Dim validSheets As Collection: Set validSheets = New Collection
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> wsSearch.Name And Not IsEmpty(GetSheetRules(ws)) Then
            validSheets.Add ws
        End If
    Next ws
    
    Dim resRow As Long: resRow = 5
    With wsSearch.Range("A5:J5")
        .Value = Array("해금 지역", "생산 시설", "필요 시설 개수", "데이터 출처", "소모 재료 관계", "회당 생산량", "생산 시간(초)", "공정 단계 (Tier)", "잔여재료", "전력소비량")
        .Font.Bold = True: .Interior.Color = RGB(220, 230, 241): .Borders.LineStyle = xlContinuous
    End With
    resRow = 6

    Dim topRow As Long, topRules As Variant
    Dim topFound As Boolean: topFound = False

    For Each ws In validSheets
        topRules = GetSheetRules(ws)
        topRow = ws.Cells(ws.Rows.Count, CStr(topRules(4))).End(xlUp).Row
        
        For i = 2 To topRow
            If ws.Cells(i, CStr(topRules(4))).MergeArea(1).Text = itemName Then
                Dim tPPS As Double: tPPS = Val(ws.Cells(i, CStr(topRules(9))).MergeArea(1).Value)
                
                If tPPS <= 0 Then
                    Dim tOut As Double: tOut = Val(ws.Cells(i, CStr(topRules(6))).MergeArea(1).Value)
                    Dim tCyc As Double: tCyc = Val(ws.Cells(i, CStr(topRules(7))).MergeArea(1).Value)
                    If tCyc <= 0 Then tCyc = 1
                    tPPS = tOut / tCyc
                End If
                
                targetPPS = tPPS
                topFound = True: Exit For
            End If
        Next i
        If topFound Then Exit For
    Next ws

    If Not topFound Then
        targetPPS = Val(wsSearch.Range("B3").Value)
        If targetPPS <= 0 Then targetPPS = 1
    End If

    ' 총 전력 소비량 계산 (G2 초기화 포함)
    Dim totalPower As Double: totalPower = 0
    wsSearch.Range("G2").Value = ""

    resRow = RunDFS(itemName, targetPPS, validSheets, resRow, wsSearch, extraQ, checkedList, True, totalPower)

    wsSearch.Range("G2").Value = totalPower

    If extraQ.Count > 0 Then
        resRow = resRow + 2
        wsSearch.Cells(resRow, 1).Value = "■ 기타 하위계보 조합법 (중복 레시피)"
        wsSearch.Range(wsSearch.Cells(resRow, 1), wsSearch.Cells(resRow, 10)).Interior.Color = RGB(240, 240, 240)
        wsSearch.Cells(resRow, 1).Font.Bold = True: resRow = resRow + 1
        
        Dim info As Variant, wsSrc As Worksheet
        For Each info In extraQ
            Set wsSrc = info(0)
            resRow = WriteDataRow(wsSrc, CLng(info(1)), CStr(info(2)), CDbl(info(3)), wsSearch, resRow, info(4), totalPower, False)
        Next info
    End If

    ApplyFullRecipeFormatting wsSearch, resRow
    wsSearch.Columns("A:J").AutoFit
    Application.ScreenUpdating = True
    MsgBox "[" & itemName & "] 최상위 1대 기준 DFS 분석 완료!"
End Sub

Function RunDFS(item As String, pps As Double, validSheets As Collection, r As Long, wsDest As Worksheet, exQ As Object, checked As Object, isPrimary As Boolean, ByRef totalPower As Double) As Long
    Dim ws As Worksheet, i As Long
    Dim candidates As Object: Set candidates = CreateObject("System.Collections.ArrayList")
    
    Dim cleanItem As String: cleanItem = Replace(item, " ", "")
    
    If Not isPrimary And checked.Exists(cleanItem) Then
        RunDFS = r: Exit Function
    End If
    If Not isPrimary Then checked.Add cleanItem, True

    For Each ws In validSheets
        Dim rules As Variant: rules = GetSheetRules(ws)
        
        Dim colMain As String: colMain = CStr(rules(4))
        Dim colSub As String: colSub = CStr(rules(5))
        Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, colMain).End(xlUp).Row
        
        For i = 2 To lastRow
            Dim isMatch As Boolean: isMatch = False
            
            If Replace(ws.Cells(i, colMain).MergeArea(1).Text, " ", "") = cleanItem Then
                isMatch = True
            ElseIf colSub <> "" Then
                If Replace(ws.Cells(i, colSub).MergeArea(1).Text, " ", "") = cleanItem Then isMatch = True
            End If
            
            If isMatch Then
                Dim facName As String: facName = ws.Cells(i, "B").MergeArea(1).Text
                Dim isLowPriority As Boolean: isLowPriority = False
                
                If InStr(facName, "휴대용") > 0 Then
                    isLowPriority = True
                ElseIf InStr(facName, "양수기") > 0 And InStr(item, "청정수") = 0 Then
                    isLowPriority = True
                End If
                
                Dim tierValue As Double
                tierValue = Val(ws.Cells(i, CStr(rules(8))).MergeArea(1).Value)
                
                candidates.Add Array(ws, i, Val(ws.Cells(i, CStr(rules(6))).MergeArea(1).Value), Val(ws.Cells(i, CStr(rules(2))).MergeArea(1).Value), rules, isLowPriority, tierValue)
            End If
        Next i
    Next ws

    SortCandidates candidates

    If candidates.Count > 0 Then
        If isPrimary Then
            wsDest.Cells(r, 1).Value = "▶ 최우선 생산 공정": wsDest.Range(wsDest.Cells(r, 1), wsDest.Cells(r, 10)).Interior.Color = RGB(255, 230, 230): r = r + 1
        End If

        Dim best As Variant: best = candidates(0)
        Dim wsBest As Worksheet: Set wsBest = best(0)
        Dim rBest As Long: rBest = CLng(best(1))
        Dim rulesBest As Variant: rulesBest = best(4)
        
        r = WriteDataRow(wsBest, rBest, item, pps, wsDest, r, rulesBest, totalPower, True)
        
        Dim sOut As Double: sOut = Val(wsBest.Cells(rBest, CStr(rulesBest(6))).MergeArea(1).Value)
        Dim cyc As Double: cyc = Val(wsBest.Cells(rBest, CStr(rulesBest(7))).MergeArea(1).Value)
        If cyc <= 0 Then cyc = 1
        
        Dim unitPPS As Double: unitPPS = Val(wsBest.Cells(rBest, CStr(rulesBest(9))).MergeArea(1).Value)
        If unitPPS <= 0 Then
            Dim tempOut As Double: tempOut = IIf(sOut <= 0, 1, sOut)
            unitPPS = tempOut / cyc
        End If
        
        Dim actualBuild As Long: actualBuild = -Int(-(pps / unitPPS - 0.000001))
        
        Dim m1 As String: m1 = Trim(wsBest.Cells(rBest, CStr(rulesBest(0))).MergeArea(1).Text)
        Dim m2 As String: m2 = ""
        If CStr(rulesBest(1)) <> "" Then m2 = Trim(wsBest.Cells(rBest, CStr(rulesBest(1))).MergeArea(1).Text)
        
        If m1 = m2 Then m2 = ""
        
        If m1 <> "" And m1 <> "-" Then
            Dim q1 As Double: q1 = Val(wsBest.Cells(rBest, CStr(rulesBest(2))).MergeArea(1).Value)
            r = RunDFS(m1, (q1 * actualBuild) / cyc, validSheets, r, wsDest, exQ, checked, False, totalPower)
        End If
        
        If m2 <> "" And m2 <> "-" Then
            Dim q2 As Double: q2 = Val(wsBest.Cells(rBest, CStr(rulesBest(3))).MergeArea(1).Value)
            r = RunDFS(m2, (q2 * actualBuild) / cyc, validSheets, r, wsDest, exQ, checked, False, totalPower)
        End If

        For i = 1 To candidates.Count - 1
            exQ.Add Array(candidates(i)(0), candidates(i)(1), item, pps, candidates(i)(4))
        Next i
    End If
    
    If Not isPrimary Then checked.Remove cleanItem
    
    RunDFS = r
End Function

Sub SortCandidates(ByRef candidates As Object)
    If candidates.Count <= 1 Then Exit Sub
    Dim j As Long, k As Long, temp As Variant
    Dim isExJ As Boolean, isExK As Boolean
    Dim tierJ As Double, tierK As Double
    
    For j = 0 To candidates.Count - 2
        For k = j + 1 To candidates.Count - 1
            Dim swap As Boolean: swap = False
            
            isExJ = candidates(j)(5)
            isExK = candidates(k)(5)
            
            If isExJ And Not isExK Then
                swap = True
            ElseIf Not isExJ And isExK Then
                swap = False
            Else
                tierJ = candidates(j)(6)
                tierK = candidates(k)(6)
                
                If tierJ > tierK Then
                    swap = True
                ElseIf tierJ = tierK Then
                    If candidates(j)(2) < candidates(k)(2) Then
                        swap = True
                    ElseIf candidates(j)(2) = candidates(k)(2) And candidates(j)(3) > candidates(k)(3) Then
                        swap = True
                    End If
                End If
            End If
            
            If swap Then
                temp = candidates(j): candidates(j) = candidates(k): candidates(k) = temp
            End If
        Next k
    Next j
End Sub

Function WriteDataRow(wsSrc As Worksheet, i As Long, target As String, pps As Double, wsDest As Worksheet, r As Long, rules As Variant, ByRef totalPower As Double, isMainLine As Boolean) As Long
    Dim sOut As Double: sOut = Val(wsSrc.Cells(i, rules(6)).MergeArea(1).Value)
    Dim cyc As Double: cyc = Val(wsSrc.Cells(i, rules(7)).MergeArea(1).Value)
    If cyc <= 0 Then cyc = 1
    
    Dim sPPS As Double: sPPS = Val(wsSrc.Cells(i, rules(9)).MergeArea(1).Value)
    If sPPS <= 0 Then
        Dim tempOut As Double: tempOut = IIf(sOut <= 0, 1, sOut)
        sPPS = tempOut / cyc
    End If
    
    Dim actualBuild As Long: actualBuild = -Int(-(pps / sPPS - 0.000001))
    
    Dim unitPower As Double
    unitPower = Val(wsSrc.Cells(i, CStr(rules(11))).MergeArea(1).Value)
    
    If isMainLine Then
        totalPower = totalPower + (actualBuild * unitPower)
    End If
    
    wsDest.Cells(r, 1).Value = wsSrc.Cells(i, 1).MergeArea(1).Value
    wsDest.Cells(r, 2).Value = wsSrc.Cells(i, "B").MergeArea(1).Value
    wsDest.Cells(r, 3).Value = actualBuild
    wsDest.Cells(r, 4).Value = wsSrc.Name
    
    Dim m1 As String: m1 = Trim(wsSrc.Cells(i, rules(0)).MergeArea(1).Text)
    Dim v1 As Double: v1 = Val(wsSrc.Cells(i, rules(2)).MergeArea(1).Value)
    Dim q1 As String: q1 = IIf(m1 = "-" Or m1 = "", "0", CStr(v1))
    
    Dim m2 As String: m2 = "": Dim q2 As String: q2 = "0"
    If CStr(rules(1)) <> "" Then
        m2 = Trim(wsSrc.Cells(i, rules(1)).MergeArea(1).Text)
        Dim v2 As Double: v2 = Val(wsSrc.Cells(i, rules(3)).MergeArea(1).Value)
        q2 = IIf(m2 = "-" Or m2 = "", "0", CStr(v2))
    End If
    
    If m1 = m2 Then
        m2 = ""
        q2 = "0"
    End If
    
    wsDest.Cells(r, 5).Value = IIf(m2 = "" Or m2 = "-", m1 & "(" & q1 & ") -> " & target & "(" & sOut & ")", m1 & "(" & q1 & ") + " & m2 & "(" & q2 & ") -> " & target & "(" & sOut & ")")
    wsDest.Cells(r, 6).Value = sOut
    wsDest.Cells(r, 7).Value = cyc
    wsDest.Cells(r, 8).Value = wsSrc.Cells(i, rules(8)).MergeArea(1).Value
    
    Dim surplus As Double: surplus = (actualBuild * sPPS) - pps
    If surplus > 0.0001 Then
        Dim surplusMin As Double: surplusMin = surplus * 60
        wsDest.Cells(r, 9).Value = target & " 분당 " & Application.Text(surplusMin, "[=0]0;[<1]0.##;#") & "개(초당 " & Application.Text(surplus, "[=0]0;[<1]0.##;#") & "개) 남음"
    End If
    
    wsDest.Cells(r, 10).Value = actualBuild * unitPower
    wsDest.Range(wsDest.Cells(r, 1), wsDest.Cells(r, 10)).Borders.LineStyle = xlContinuous
    WriteDataRow = r + 1
End Function

Function GetSheetRules(ws As Worksheet) As Variant
    Dim zVal As Variant
    zVal = ws.Range("Z1").Value
    
    If IsError(zVal) Or IsEmpty(zVal) Then
        GetSheetRules = Empty
        Exit Function
    End If
    
    If CStr(zVal) = "2" Then
        GetSheetRules = Array("C", "D", "E", "F", "G", "H", "K", "J", "I", "L", "M", "N")
    ElseIf CStr(zVal) = "1" Then
        GetSheetRules = Array("C", "", "D", "", "E", "", "H", "G", "F", "I", "J", "K")
    Else
        GetSheetRules = Empty
    End If
End Function

Private Sub InitializeSearchSheet(ws As Worksheet)
    Dim lastR As Long: lastR = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row: If lastR < 5 Then lastR = 500
    ws.Range("K:Z").Clear
    With ws.Rows("5:" & lastR + 500)
        .UnMerge: .ClearContents: .Borders.LineStyle = xlNone: .Interior.ColorIndex = xlNone: .FormatConditions.Delete
    End With
End Sub

Private Sub ApplyFullRecipeFormatting(ws As Worksheet, lastRow As Long)
    Dim rngA As Range: Set rngA = ws.Range("A6:A" & lastRow)
    Dim rngC As Range: Set rngC = ws.Range("C6:C" & lastRow)
    Dim rngH As Range: Set rngH = ws.Range("H6:H" & lastRow)
    If lastRow >= 6 Then
        With rngA.FormatConditions.Add(Type:=xlTextString, String:="무릉", TextOperator:=xlContains): .Interior.Color = RGB(184, 222, 205): End With
        With rngA.FormatConditions.Add(Type:=xlTextString, String:="4번 협곡", TextOperator:=xlContains): .Interior.Color = RGB(240, 240, 172): End With
        
        Dim colRng As Variant: For Each colRng In Array(rngC, rngH)
            With colRng.FormatConditions.AddDatabar: .ShowValue = True: .BarFillType = xlDataBarFillSolid: .BarColor.Color = RGB(203, 205, 255): End With
        Next colRng
    End If
End Sub
