Option Explicit

' 텍스트 내 보이지 않는 줄바꿈, 공백을 완벽히 제거하여 자원 매칭 오류를 막는 함수
Function CleanStr(ByVal s As String) As String
    If IsNull(s) Then s = ""
    s = Replace(s, " ", "")
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    s = Replace(s, Chr(160), "")
    CleanStr = s
End Function

' ==========================================
' 1. 기본 검색 함수 (D3 셀 기준)
' ==========================================
Sub Search()
    Dim wsSearch As Worksheet, ws As Worksheet
    Dim searchWord As String
    Dim lastRow As Long, i As Long, j As Long, resRow As Long
    Dim foundCount As Integer
    Dim rules As Variant
    
    Set wsSearch = ThisWorkbook.sheets("검색")
    searchWord = wsSearch.Range("D3").Value
    
    If searchWord = "" Then
        MsgBox "검색어를 입력해주세요."
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' 6행부터 완전 초기화 (서식 포함)
    wsSearch.Rows("6:" & wsSearch.Rows.Count).Clear
    wsSearch.Range("G3").Value = ""
    wsSearch.Range("L3").Value = ""
    wsSearch.Range("A5:B500").Clear
    
    resRow = 6

    ' [생산 품목] 검색 - 첫 번째(메인)와 두 번째(부산물) 생산품 모두 확인
    wsSearch.Cells(resRow, 1).Value = "▶ [" & searchWord & "]을(를) 생산하는 공정"
    wsSearch.Cells(resRow, 1).Font.Bold = True
    resRow = resRow + 1
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> wsSearch.Name Then
            rules = GetSheetRules(ws)
            If IsArray(rules) Then
                lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
                foundCount = 0
                For i = 2 To lastRow
                    Dim out1 As String: out1 = ws.Cells(i, CStr(rules(4))).MergeArea(1).Value
                    Dim out2 As String: out2 = ""
                    If CStr(rules(5)) <> "" Then out2 = ws.Cells(i, CStr(rules(5))).MergeArea(1).Value
                    
                    If InStr(out1, searchWord) > 0 Or (out2 <> "" And InStr(out2, searchWord) > 0) Then
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
    wsSearch.Cells(resRow, 1).Value = "▶ [" & searchWord & "]을(를) 재료로 소모하는 공정"
    wsSearch.Cells(resRow, 1).Font.Bold = True
    resRow = resRow + 1
    
    ' [소모 품목] 검색 - 첫 번째 재료와 두 번째 재료 모두 확인
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> wsSearch.Name Then
            rules = GetSheetRules(ws)
            If IsArray(rules) Then
                lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
                foundCount = 0
                For i = 2 To lastRow
                    Dim in1 As String: in1 = ws.Cells(i, CStr(rules(0))).MergeArea(1).Value
                    Dim in2 As String: in2 = ""
                    If CStr(rules(1)) <> "" Then in2 = ws.Cells(i, CStr(rules(1))).MergeArea(1).Value
                    
                    If InStr(in1, searchWord) > 0 Or (in2 <> "" And InStr(in2, searchWord) > 0) Then
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

    ' 기본 검색 결과 중앙 정렬
    If resRow > 6 Then
        wsSearch.Range("A6:O" & resRow).HorizontalAlignment = xlCenter
    End If

    Application.CutCopyMode = False
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    wsSearch.Columns.AutoFit
    MsgBox "검색 완료"
End Sub

' ==========================================
' 2. 상세 레시피 DFS 분석 함수
' ==========================================
Sub SearchFullRecipe()
    Dim wsSearch As Worksheet: Set wsSearch = ThisWorkbook.sheets("검색")
    Dim itemName As String
    Dim ws As Worksheet
    Dim i As Long
    
    itemName = Trim(wsSearch.Range("G3").Value)
    If itemName = "" Then
        itemName = Trim(Selection.Cells(1, 1).MergeArea(1).Text)
        If itemName = "" Then
            MsgBox "아이템 명을 선택하거나 G3 셀에 입력해주세요."
            Exit Sub
        End If
        wsSearch.Range("G3").Value = itemName
    End If

    Application.ScreenUpdating = False
    InitializeSearchSheet wsSearch

    Dim targetPPS As Double
    Dim extraQ As Object: Set extraQ = CreateObject("System.Collections.ArrayList")
    Dim checkedList As Object: Set checkedList = CreateObject("Scripting.Dictionary")
    Dim dictLiquids As Object: Set dictLiquids = CreateObject("Scripting.Dictionary")
    Dim dictSolids As Object: Set dictSolids = CreateObject("Scripting.Dictionary")
    
    Dim validSheets As Collection: Set validSheets = New Collection
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> wsSearch.Name And Not IsEmpty(GetSheetRules(ws)) Then
            validSheets.Add ws
            If ws.Name = "자원 채집" Then
                Dim tRules As Variant: tRules = GetSheetRules(ws)
                Dim lr As Long: lr = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
                For i = 2 To lr
                    Dim fName As String: fName = ws.Cells(i, "B").MergeArea(1).Text
                    Dim oName1 As String: oName1 = CleanStr(ws.Cells(i, CStr(tRules(4))).MergeArea(1).Text)
                    Dim oName2 As String: oName2 = ""
                    If CStr(tRules(5)) <> "" Then oName2 = CleanStr(ws.Cells(i, CStr(tRules(5))).MergeArea(1).Text)
                    
                    If InStr(fName, "양수기") > 0 Or InStr(fName, "내산성 양수기") > 0 Then
                        If oName1 <> "" And oName1 <> "-" Then dictLiquids(oName1) = 0
                        If oName2 <> "" And oName2 <> "-" Then dictLiquids(oName2) = 0
                    ElseIf InStr(fName, "채굴기") > 0 Then
                        If oName1 <> "" And oName1 <> "-" Then dictSolids(oName1) = 0
                        If oName2 <> "" And oName2 <> "-" Then dictSolids(oName2) = 0
                    End If
                Next i
            End If
        End If
    Next ws
    
    dictLiquids("오염수") = 0
    
    Dim resRow As Long: resRow = 6

    Dim topFound As Boolean: topFound = False
    For Each ws In validSheets
        Dim topRules As Variant: topRules = GetSheetRules(ws)
        Dim topRow As Long: topRow = ws.Cells(ws.Rows.Count, CStr(topRules(4))).End(xlUp).Row
        For i = 2 To topRow
            Dim topOut1 As String: topOut1 = CleanStr(ws.Cells(i, CStr(topRules(4))).MergeArea(1).Text)
            Dim topOut2 As String: topOut2 = ""
            If CStr(topRules(5)) <> "" Then topOut2 = CleanStr(ws.Cells(i, CStr(topRules(5))).MergeArea(1).Text)
            
            If topOut1 = CleanStr(itemName) Or topOut2 = CleanStr(itemName) Then
                Dim tPPS As Double: tPPS = Val(ws.Cells(i, CStr(topRules(9))).MergeArea(1).Value)
                If tPPS <= 0 Then tPPS = Val(ws.Cells(i, CStr(topRules(10))).MergeArea(1).Value) / 60
                If tPPS <= 0 Then
                    Dim tOut As Double: tOut = Val(ws.Cells(i, CStr(topRules(6))).MergeArea(1).Value)
                    Dim tCyc As Double: tCyc = Val(ws.Cells(i, CStr(topRules(7))).MergeArea(1).Value)
                    If tCyc <= 0 Then tCyc = 1
                    tPPS = tOut / tCyc
                End If
                targetPPS = tPPS: topFound = True: Exit For
            End If
        Next i
        If topFound Then Exit For
    Next ws
    If Not topFound Then targetPPS = 1

    Dim utilRate As Double: utilRate = Val(wsSearch.Range("J3").Value)
    If utilRate <= 0 Or utilRate > 1 Then utilRate = 1

    Dim totalPower As Double: totalPower = 0
    resRow = RunDFS(itemName, targetPPS * utilRate, validSheets, resRow, wsSearch, extraQ, checkedList, True, totalPower, dictLiquids, dictSolids)

    wsSearch.Range("L3").Value = Round(totalPower, 1)

    If extraQ.Count > 0 Then
        resRow = resRow + 2
        wsSearch.Cells(resRow, 3).Value = "■ 기타 하위계보 조합법 (중복 레시피)"
        wsSearch.Range(wsSearch.Cells(resRow, 3), wsSearch.Cells(resRow, 13)).Interior.Color = RGB(240, 240, 240)
        wsSearch.Cells(resRow, 3).Font.Bold = True: resRow = resRow + 1
        Dim info As Variant, wsSrc As Worksheet
        For Each info In extraQ
            Set wsSrc = info(0)
            resRow = WriteDataRow(wsSrc, CLng(info(1)), CStr(info(2)), CDbl(info(3)), wsSearch, resRow, info(4), totalPower, False)
        Next info
    End If

    ApplyFullRecipeFormatting wsSearch, resRow
    DrawResourceSummary wsSearch, dictLiquids, dictSolids
    wsSearch.Columns("A:M").AutoFit
    Application.ScreenUpdating = True
    MsgBox "[" & itemName & "] 레시피 분석 완료!"
End Sub

' --- 자원 요약표 ---
Private Sub DrawResourceSummary(ByVal ws As Worksheet, ByVal dictLiquids As Object, ByVal dictSolids As Object)
    Dim rowIdx As Long
    Dim k As Variant
    Dim hasLiquid As Boolean, hasSolid As Boolean
    
    hasLiquid = False
    For Each k In dictLiquids.keys
        If dictLiquids(k) > 0 Then
            hasLiquid = True
            Exit For
        End If
    Next k
    
    hasSolid = False
    For Each k In dictSolids.keys
        If dictSolids(k) > 0 Then
            hasSolid = True
            Exit For
        End If
    Next k

    rowIdx = 5
    
    ws.Range("A5:B50").Clear
    ws.Range("A5:B50").Borders.LineStyle = xlNone
    
    ws.Range("A" & rowIdx).Value = "필요한 자원"
    ws.Range("B" & rowIdx).Value = "분당 생산량"
    ws.Range("A" & rowIdx & ":B" & rowIdx).Interior.Color = RGB(252, 228, 214)
    ws.Range("A" & rowIdx & ":B" & rowIdx).Font.Bold = True
    rowIdx = rowIdx + 1
    
    If hasLiquid Then
        ws.Range("A" & rowIdx).Value = "액체"
        ws.Range("A" & rowIdx & ":B" & rowIdx).Merge
        ws.Range("A" & rowIdx).Interior.Color = RGB(221, 235, 247)
        rowIdx = rowIdx + 1
        
        ws.Range("A" & rowIdx).Value = "파이프 속도"
        ws.Range("B" & rowIdx).Value = "120/m"
        rowIdx = rowIdx + 1
        
        For Each k In dictLiquids.keys
            If dictLiquids(k) > 0 Then
                ws.Range("A" & rowIdx).Value = k
                ws.Range("B" & rowIdx).Value = dictLiquids(k)
                rowIdx = rowIdx + 1
            End If
        Next k
    End If
    
    If hasSolid Then
        ws.Range("A" & rowIdx).Value = "고체"
        ws.Range("A" & rowIdx & ":B" & rowIdx).Merge
        ws.Range("A" & rowIdx).Interior.Color = RGB(255, 242, 204)
        rowIdx = rowIdx + 1
        
        ws.Range("A" & rowIdx).Value = "컨베이어 벨트 속도"
        ws.Range("B" & rowIdx).Value = "30/m"
        rowIdx = rowIdx + 1
        
        For Each k In dictSolids.keys
            If dictSolids(k) > 0 Then
                ws.Range("A" & rowIdx).Value = k
                ws.Range("B" & rowIdx).Value = dictSolids(k)
                rowIdx = rowIdx + 1
            End If
        Next k
    End If
    
    If rowIdx > 6 Then
        ws.Range("A5:B" & (rowIdx - 1)).Borders.LineStyle = xlContinuous
        ws.Range("A5:B" & (rowIdx - 1)).HorizontalAlignment = xlCenter
    End If
End Sub

' 공정 티어 탐색
Function GetItemTier(ByVal item As String, ByVal validSheets As Collection) As Double
    Dim cleanItem As String: cleanItem = CleanStr(item)
    Dim ws As Worksheet, i As Long, minTier As Double: minTier = 999: Dim found As Boolean: found = False
    For Each ws In validSheets
        Dim rules As Variant: rules = GetSheetRules(ws)
        Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, CStr(rules(4))).End(xlUp).Row
        For i = 2 To lastRow
            Dim tOut1 As String: tOut1 = CleanStr(ws.Cells(i, CStr(rules(4))).MergeArea(1).Text)
            Dim tOut2 As String: tOut2 = ""
            If CStr(rules(5)) <> "" Then tOut2 = CleanStr(ws.Cells(i, CStr(rules(5))).MergeArea(1).Text)
            
            If tOut1 = cleanItem Or tOut2 = cleanItem Then
                Dim t As Double: t = Val(ws.Cells(i, CStr(rules(8))).MergeArea(1).Value)
                If t < minTier Then minTier = t: found = True
            End If
        Next i
    Next ws
    If found Then GetItemTier = minTier Else GetItemTier = 0
End Function

' DFS 탐색
Function RunDFS(ByVal item As String, ByVal pps As Double, ByVal validSheets As Collection, ByVal r As Long, ByVal wsDest As Worksheet, ByVal exQ As Object, ByVal checked As Object, ByVal isPrimary As Boolean, ByRef totalPower As Double, ByVal dictLiquids As Object, ByVal dictSolids As Object, Optional ByVal blockDepth As Integer = 1) As Long
    Dim ws As Worksheet, i As Long, cleanItem As String: cleanItem = CleanStr(item)
    If Not isPrimary Then
        If dictLiquids.Exists(cleanItem) Then dictLiquids(cleanItem) = dictLiquids(cleanItem) + (pps * 60)
        If dictSolids.Exists(cleanItem) Then dictSolids(cleanItem) = dictSolids(cleanItem) + (pps * 60)
    End If
    If Not isPrimary And checked.Exists(cleanItem) Then RunDFS = r: Exit Function
    If Not isPrimary Then checked.Add cleanItem, True

    Dim candidates As Object: Set candidates = CreateObject("System.Collections.ArrayList")
    For Each ws In validSheets
        Dim rules As Variant: rules = GetSheetRules(ws)
        Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, CStr(rules(4))).End(xlUp).Row
        For i = 2 To lastRow
            Dim cOut1 As String: cOut1 = CleanStr(ws.Cells(i, CStr(rules(4))).MergeArea(1).Text)
            Dim cOut2 As String: cOut2 = ""
            If CStr(rules(5)) <> "" Then cOut2 = CleanStr(ws.Cells(i, CStr(rules(5))).MergeArea(1).Text)
            
            If cOut1 = cleanItem Or cOut2 = cleanItem Then
                Dim facName As String: facName = ws.Cells(i, "B").MergeArea(1).Text
                Dim isLow As Boolean: isLow = False
                If InStr(facName, "휴대용") > 0 Or ((InStr(facName, "양수기") > 0 Or InStr(facName, "내산성 양수기 II") > 0) And (InStr(item, "청정수") = 0 And InStr(item, "산성 침적물") = 0)) Then isLow = True
                
                Dim isExt As Boolean: isExt = (InStr(facName, "확장") > 0)
                candidates.Add Array(ws, i, Val(ws.Cells(i, CStr(rules(6))).MergeArea(1).Value), Val(ws.Cells(i, CStr(rules(2))).MergeArea(1).Value), rules, isLow, Val(ws.Cells(i, CStr(rules(8))).MergeArea(1).Value), isExt)
            End If
        Next i
    Next ws
    SortCandidates candidates

    If candidates.Count > 0 Then
        If isPrimary Then
            wsDest.Cells(r, 3).Value = "▶ 최우선 생산 공정": wsDest.Range(wsDest.Cells(r, 3), wsDest.Cells(r, 13)).Interior.Color = RGB(255, 230, 230): r = r + 1
        End If
        Dim best As Variant: best = candidates(0)
        r = WriteDataRow(best(0), CLng(best(1)), item, pps, wsDest, r, best(4), totalPower, True)
        
        Dim sOut As Double: sOut = Val(best(0).Cells(best(1), best(4)(6)).MergeArea(1).Value): Dim tOutVal As Double: tOutVal = IIf(sOut <= 0, 1, sOut)
        
        Dim m1 As String: m1 = Trim(best(0).Cells(best(1), best(4)(0)).MergeArea(1).Text)
        Dim v1_val As Double: v1_val = Val(best(0).Cells(best(1), best(4)(2)).MergeArea(1).Value)
        
        Dim m2 As String: m2 = ""
        Dim v2_val As Double: v2_val = 0
        
        If CStr(best(4)(1)) <> "" Then
            m2 = Trim(best(0).Cells(best(1), best(4)(1)).MergeArea(1).Text)
            v2_val = Val(best(0).Cells(best(1), best(4)(3)).MergeArea(1).Value)
        End If
        
        ' ==============================================================
        ' [수정] 병합된 셀 처리: 이름이 같으면 두 번째 재료는 무시 (합산 아님)
        ' ==============================================================
        If m1 = m2 Then
            m2 = ""
            v2_val = 0
        End If
        ' ==============================================================
        
        Dim t1 As Double: t1 = 0: Dim t2 As Double: t2 = 0
        If m1 <> "" And m1 <> "-" Then t1 = GetItemTier(m1, validSheets)
        If m2 <> "" And m2 <> "-" Then t2 = GetItemTier(m2, validSheets)
        Dim split As Boolean: split = (t1 >= 5 And t2 >= 5): Dim nextDepth As Integer: nextDepth = IIf(split, blockDepth + 1, blockDepth)
        
        If m1 <> "" And m1 <> "-" Then
            If split Then
                wsDest.Cells(r, 7).Value = String(blockDepth, "v") & " " & m1 & " 생성 공정": wsDest.Cells(r, 7).HorizontalAlignment = xlCenter: wsDest.Cells(r, 7).Font.Bold = True: wsDest.Range(wsDest.Cells(r, 3), wsDest.Cells(r, 13)).Interior.Color = RGB(226, 239, 218): wsDest.Range(wsDest.Cells(r, 3), wsDest.Cells(r, 13)).Borders.LineStyle = xlContinuous: r = r + 1
            End If
            r = RunDFS(m1, pps * (v1_val / tOutVal), validSheets, r, wsDest, exQ, checked, False, totalPower, dictLiquids, dictSolids, nextDepth)
            If split Then
                wsDest.Cells(r, 7).Value = String(blockDepth, "^") & " " & m1 & " 생성 공정": wsDest.Cells(r, 7).HorizontalAlignment = xlCenter: wsDest.Cells(r, 7).Font.Bold = True: wsDest.Range(wsDest.Cells(r, 3), wsDest.Cells(r, 13)).Interior.Color = RGB(226, 239, 218): wsDest.Range(wsDest.Cells(r, 3), wsDest.Cells(r, 13)).Borders.LineStyle = xlContinuous: r = r + 1
            End If
        End If
        If m2 <> "" And m2 <> "-" Then
            If split Then
                wsDest.Cells(r, 7).Value = String(blockDepth, "v") & " " & m2 & " 생성 공정": wsDest.Cells(r, 7).HorizontalAlignment = xlCenter: wsDest.Cells(r, 7).Font.Bold = True: wsDest.Range(wsDest.Cells(r, 3), wsDest.Cells(r, 13)).Interior.Color = RGB(226, 239, 218): wsDest.Range(wsDest.Cells(r, 3), wsDest.Cells(r, 13)).Borders.LineStyle = xlContinuous: r = r + 1
            End If
            r = RunDFS(m2, pps * (v2_val / tOutVal), validSheets, r, wsDest, exQ, checked, False, totalPower, dictLiquids, dictSolids, nextDepth)
            If split Then
                wsDest.Cells(r, 7).Value = String(blockDepth, "^") & " " & m2 & " 생성 공정": wsDest.Cells(r, 7).HorizontalAlignment = xlCenter: wsDest.Cells(r, 7).Font.Bold = True: wsDest.Range(wsDest.Cells(r, 3), wsDest.Cells(r, 13)).Interior.Color = RGB(226, 239, 218): wsDest.Range(wsDest.Cells(r, 3), wsDest.Cells(r, 13)).Borders.LineStyle = xlContinuous: r = r + 1
            End If
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
    Dim j As Long, k As Long, temp As Variant, swap As Boolean
    For j = 0 To candidates.Count - 2
        For k = j + 1 To candidates.Count - 1
            swap = False
            If candidates(j)(5) And Not candidates(k)(5) Then
                swap = True
            ElseIf Not candidates(j)(5) And Not candidates(k)(5) Then
                If candidates(j)(6) > candidates(k)(6) Then
                    swap = True
                ElseIf candidates(j)(6) = candidates(k)(6) Then
                    If Not candidates(j)(7) And candidates(k)(7) Then
                        swap = True
                    ElseIf candidates(j)(7) = candidates(k)(7) And candidates(j)(2) < candidates(k)(2) Then
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

Function WriteDataRow(ByVal wsSrc As Worksheet, ByVal i As Long, ByVal target As String, ByVal pps As Double, ByVal wsDest As Worksheet, ByVal r As Long, ByVal rules As Variant, ByRef totalPower As Double, ByVal isMain As Boolean) As Long
    ' 1. 기본 데이터 추출
    Dim sOutCount As Double: sOutCount = Val(wsSrc.Cells(i, rules(6)).MergeArea(1).Value) ' 회당 생산량
    If sOutCount <= 0 Then sOutCount = 1
    
    Dim sTime As Double: sTime = Val(wsSrc.Cells(i, rules(7)).MergeArea(1).Value) ' 생산 시간(초)
    If sTime <= 0 Then sTime = 1
    
    ' 2. 시설의 순수 생산 성능 (PPS) 계산
    Dim sPPS As Double: sPPS = sOutCount / sTime
    
    ' 3. 시설명 및 기본 전력 추출
    Dim facName As String: facName = wsSrc.Cells(i, "B").MergeArea(1).Text
    Dim unitP As Double: unitP = Val(wsSrc.Cells(i, CStr(rules(11))).MergeArea(1).Value)
    
    ' [핵심 수정] 정확한 필요 가동 공정 수(exact) 계산
    ' pps(목표 초당 생산량)를 시설의 초당 생산 성능(sPPS)으로 나눕니다.
    Dim exact As Double: exact = pps / sPPS
    
    ' 4. 반응기 -> 확장 반응기 자동 변환 로직
    If facName = "반응기" And exact > 1 Then
        facName = "확장 반응기"
        unitP = 50
    End If
    
    Dim isExt As Boolean: isExt = (InStr(facName, "확장") > 0)
    
    ' 5. 실제 물리적 기계 대수 (buildNum) 계산 (확장은 3공정당 1대)
    Dim buildNum As Double: buildNum = IIf(isExt, exact / 3, exact)
    
    ' 6. 전체 전력 소비량 누적 (가동 공정 수 * 단위 전력)
    ' exact가 1.0(1대분)이면 딱 20W(재배기)가 나오게 됩니다.
    If isMain Then totalPower = totalPower + (exact * unitP)
    
    ' 7. 결과 시트 작성
    wsDest.Cells(r, 3).Value = wsSrc.Cells(i, 1).MergeArea(1).Value ' 해금 지역
    wsDest.Cells(r, 4).Value = facName ' 시설명
    
    ' 필요 시설 개수 및 공정 수 텍스트 정리 (1.0 -> 1 처리)
    Dim txtBuild As String: txtBuild = Replace(Application.Text(buildNum, "0.##") & "대", ".대", "대")
    Dim txtExact As String: txtExact = Replace(Application.Text(exact, "0.##") & "공정", ".공정", "공정")
    wsDest.Cells(r, 5).Value = txtBuild & IIf(isExt, " (" & txtExact & ")", "")
    
    wsDest.Cells(r, 6).Value = wsSrc.Name ' 데이터 출처
    wsDest.Cells(r, 8).Value = sOutCount ' 회당 생산량
    wsDest.Cells(r, 9).Value = sTime ' 생산 시간
    wsDest.Cells(r, 10).Value = wsSrc.Cells(i, rules(8)).MergeArea(1).Value ' 티어
    
    ' 소모 재료 관계 작성
    Dim m1 As String: m1 = Trim(wsSrc.Cells(i, rules(0)).MergeArea(1).Text)
    Dim v1 As Double: v1 = Val(wsSrc.Cells(i, rules(2)).MergeArea(1).Value)
    Dim m2 As String: m2 = ""
    Dim v2 As Double: v2 = 0
    If CStr(rules(1)) <> "" Then
        m2 = Trim(wsSrc.Cells(i, rules(1)).MergeArea(1).Text)
        v2 = Val(wsSrc.Cells(i, rules(3)).MergeArea(1).Value)
    End If
    
    If m1 = m2 Then m2 = "": v2 = 0
    wsDest.Cells(r, 7).Value = IIf(m2 = "" Or m2 = "-", m1 & "(" & v1 & ") -> " & target, m1 & "(" & v1 & ") + " & m2 & "(" & v2 & ") -> " & target)
    
    ' 흑자 계산
    Dim ceilB As Long: ceilB = Application.WorksheetFunction.RoundUp(buildNum, 0)
    Dim surplus As Double: surplus = (ceilB * IIf(isExt, 3, 1) * sPPS) - pps
    wsDest.Cells(r, 11).Value = IIf(surplus > 0.0001, "분당 " & Application.Text(surplus * 60, "0.##") & "개 흑자", "-")
    
    ' [최종 수정] 개별 전력 소비량 출력 (L열)
    wsDest.Cells(r, 12).Value = Application.WorksheetFunction.Round(exact * unitP, 1)
    
    wsDest.Range(wsDest.Cells(r, 3), wsDest.Cells(r, 13)).Borders.LineStyle = xlContinuous
    WriteDataRow = r + 1
End Function
Function GetSheetRules(ByVal ws As Worksheet) As Variant
    Dim zVal As Variant: zVal = ws.Range("Z1").Value
    If IsError(zVal) Or IsEmpty(zVal) Then: GetSheetRules = Empty: Exit Function
    If CStr(zVal) = "2" Then GetSheetRules = Array("C", "D", "E", "F", "G", "H", "K", "J", "I", "L", "M", "N")
    If CStr(zVal) = "1" Then GetSheetRules = Array("C", "", "D", "", "E", "", "H", "G", "F", "I", "J", "K")
End Function

Private Sub InitializeSearchSheet(ByVal ws As Worksheet)
    Dim lastR As Long
    lastR = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    If lastR < 6 Then lastR = 500
    Dim lastRA As Long
    lastRA = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRA > lastR Then lastR = lastRA

    ws.Range("A5:B" & lastRA + 500).Clear
    ws.Range("N:Z").Clear
    ws.Rows("6:" & lastR + 500).Clear
    
    With ws.Range("C1:M1")
        .Value = Array("해금 지역", "생산 시설", "필요 시설 개수", "데이터 출처", "소모 재료 관계", "회당 생산량", "생산 시간(초)", "공정 단계 (Tier)", "잔여재료", "전력소비량", "분당 생산량")
        .Font.Bold = True
        .Interior.Color = RGB(220, 230, 241)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
    End With
End Sub

Private Sub ApplyFullRecipeFormatting(ByVal ws As Worksheet, ByVal lastRow As Long)
    If lastRow < 6 Then Exit Sub
    
    ' 모든 텍스트 중앙 정렬
    ws.Range("C6:M" & lastRow).HorizontalAlignment = xlCenter
    
    Dim rngE As Range: Set rngE = ws.Range("E6:E" & lastRow)
    Dim rngJ As Range: Set rngJ = ws.Range("J6:J" & lastRow)
    Dim rngC As Range: Set rngC = ws.Range("C6:C" & lastRow)
    Dim rngK As Range: Set rngK = ws.Range("K6:K" & lastRow)
    Dim rngM As Range: Set rngM = ws.Range("M6:M" & lastRow)
    
    ' 1. 잔여 재료 텍스트 다듬기
    rngK.Replace What:=".개", Replacement:="개", LookAt:=xlPart
    
    ' 2. 확장 시설 관리용 변수 (Dictionary 사용)
    Dim facDict As Object: Set facDict = CreateObject("Scripting.Dictionary")
    Dim cell As Range, rIdx As Long
    
    For Each cell In rngE
        rIdx = cell.Row
        If Not IsEmpty(cell.Value) Then
            Dim strVal As String: strVal = CStr(cell.Value)
            Dim valNum As Double: valNum = Val(strVal)
            Dim facBaseName As String: facBaseName = CStr(ws.Range("D" & rIdx).Value)
            Dim recipeKey As String: recipeKey = CleanStr(ws.Range("G" & rIdx).Value) ' 레시피 고유 식별자
            
            ' 확장 시설인 경우
            If InStr(facBaseName, "확장") > 0 Then
                Dim exactProc As Double ' [수정] 소수점 공정을 날리지 않도록 Double로 변경
                If InStr(strVal, "(") > 0 Then
                    exactProc = Val(split(split(strVal, "(")(1), "공정")(0))
                Else
                    exactProc = 1
                End If
                
                Dim neededAmount As Double: neededAmount = exactProc
                
                ' [수정] A~A 중복 방지를 위해 매 행마다 컬렉션 초기화
                Dim usedLabels As Collection: Set usedLabels = New Collection
                
                ' 필요한 가동률이 0이 될 때까지 시설(A, B, C...) 탐색 및 배정
                Do While neededAmount > 0.0001
                    Dim tIdx As Integer: tIdx = 0
                    Dim allocated As Boolean: allocated = False
                    
                    ' 적절한 타워 찾기 (A부터 시작)
                    Do While tIdx < 26
                        Dim tKey_Slots As String: tKey_Slots = facBaseName & "_" & tIdx & "_slots"
                        Dim tKey_Recipes As String: tKey_Recipes = facBaseName & "_" & tIdx & "_recipes"
                        
                        If Not facDict.Exists(tKey_Slots) Then
                            facDict(tKey_Slots) = 0#
                            facDict(tKey_Recipes) = "|"
                        End If
                        
                        Dim availableRoom As Double: availableRoom = 3# - facDict(tKey_Slots)
                        
                        ' 조건: 슬롯이 남아있고 (최대 3), 해당 타워에 이 레시피가 중복되지 않아야 함
                        If availableRoom > 0.0001 And InStr(facDict(tKey_Recipes), "|" & recipeKey & "|") = 0 Then
                            ' 현재 타워에 배정할 양 계산 (1개 레시피는 한 타워에서 최대 1.0(1대)까지만 가동 가능)
                            Dim putLimit As Double: putLimit = 1#
                            If availableRoom < putLimit Then putLimit = availableRoom
                            If neededAmount < putLimit Then putLimit = neededAmount
                            
                            ' 데이터 업데이트
                            facDict(tKey_Slots) = facDict(tKey_Slots) + putLimit
                            facDict(tKey_Recipes) = facDict(tKey_Recipes) & recipeKey & "|"
                            
                            ' 레이블 추가 (A, B, C...)
                            Dim lbl As String: lbl = Chr(65 + tIdx)
                            Dim alreadyHas As Boolean: alreadyHas = False
                            Dim u As Variant
                            For Each u In usedLabels
                                If u = lbl Then alreadyHas = True: Exit For
                            Next u
                            If Not alreadyHas Then usedLabels.Add lbl
                            
                            neededAmount = neededAmount - putLimit
                            allocated = True
                            Exit Do
                        End If
                        tIdx = tIdx + 1
                    Loop
                    
                    If Not allocated Then Exit Do ' 방어 로직: 모든 타워가 가득 참
                Loop
                
                ' 시설 이름 업데이트 (예: 확장 반응기 A~B)
                Dim finalLabel As String
                If usedLabels.Count = 1 Then
                    finalLabel = usedLabels(1)
                ElseIf usedLabels.Count > 1 Then
                    finalLabel = usedLabels(1) & "~" & usedLabels(usedLabels.Count)
                End If
                ws.Range("D" & rIdx).Value = facBaseName & " " & finalLabel
                
                ' 필요 시설 개수 텍스트 정리
                cell.Value = Replace(Application.Text(valNum, "0.##") & "대", ".대", "대")
                If exactProc > 0 Then
                    cell.Value = cell.Value & " (" & Replace(Application.Text(exactProc, "0.##") & "공정", ".공정", "공정") & ")"
                End If
                
            Else
                ' 일반 시설: 기존처럼 올림 처리하여 '정수 대수'로 표시
                If valNum > 0 Then
                    Dim roundedNum As Double: roundedNum = Application.WorksheetFunction.RoundUp(valNum, 0)
                    If IsNumeric(strVal) Then
                        cell.Value = roundedNum & "대"
                    Else
                        cell.Value = Replace(strVal, CStr(valNum), CStr(roundedNum), 1, 1)
                    End If
                End If
                cell.Value = Replace(cell.Value, ".대", "대")
            End If
        End If
    Next cell
    
    ' 3. 고정값 세팅
    ws.Range("K6").Value = ""
    ws.Range("K7").Value = "-"
    
    ' ==============================================================
    ' [수정] 4. M열(분당 생산량) 계산 로직:
    ' 확장 시설의 경우 기계 대수(0.17 등)가 아닌 실제 가동 공정 수(0.5 등)를
    ' 기준으로 계산하여 소수점 오류를 완벽히 제거
    ' ==============================================================
    Dim i As Long
    For i = 6 To lastRow
        If IsNumeric(ws.Range("H" & i).Value) And IsNumeric(ws.Range("I" & i).Value) Then
            If ws.Range("I" & i).Value > 0 Then
                Dim facN As String: facN = CStr(ws.Range("D" & i).Value)
                Dim eStr As String: eStr = CStr(ws.Range("E" & i).Value)
                Dim processCount As Double
                
                If InStr(facN, "확장") > 0 Then
                    ' E열에서 ( ) 안의 공정 수를 정확히 추출
                    If InStr(eStr, "(") > 0 Then
                        processCount = Val(split(split(eStr, "(")(1), "공정")(0))
                    Else
                        processCount = Val(eStr) * 3 ' 예외 처리
                    End If
                Else
                    processCount = Val(eStr)
                End If
                
                If processCount > 0 Then
                    Dim basePPM As Double: basePPM = (ws.Range("H" & i).Value / ws.Range("I" & i).Value) * 60
                    Dim actualPPS As Double: actualPPS = basePPM * processCount
                    ws.Range("M" & i).Value = Application.WorksheetFunction.Round(actualPPS, 2)
                End If
            End If
        End If
    Next i
    
    ' M열 숫자 서식
    With rngM: .NumberFormat = "General": End With
    
    ' 5. 조건부 서식 (무릉/협곡 등 기존 로직 유지)
    With rngC.FormatConditions.Add(Type:=xlTextString, String:="무릉", TextOperator:=xlContains)
        .Interior.Color = RGB(184, 222, 205)
    End With
    With rngC.FormatConditions.Add(Type:=xlTextString, String:="4번 협곡", TextOperator:=xlContains)
        .Interior.Color = RGB(240, 240, 172)
    End With
    
    ' 데이터 막대 표시
    Dim col As Variant
    For Each col In Array(rngE, rngJ)
        With col.FormatConditions.AddDatabar
            .ShowValue = True
            .BarFillType = xlDataBarFillSolid
            .BarColor.Color = RGB(203, 205, 255)
        End With
    Next col
End Sub
