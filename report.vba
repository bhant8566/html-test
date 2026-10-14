'==========================
' Module: modRtuHtmlReport
'==========================
Option Explicit
' -------------------------
' Helpers
' -------------------------
Private Function SanitizeOneLine_PT(ByVal s As String) As String
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    SanitizeOneLine_PT = Trim$(s)
End Function

Private Function ResolveAlarmTxtForHtml_(ByVal originalTxt As String, ByVal webTxt As String, ByVal wendyTxt As String) As String
    Dim rawTxt As String
    Dim webAlarmTxt As String, wendyAlarmTxt As String

    rawTxt = CStr(originalTxt)
    webAlarmTxt = Trim$(CStr(webTxt))
    wendyAlarmTxt = Trim$(CStr(wendyTxt))

    If checkTrim(rawTxt) And Len(webAlarmTxt) > 0 Then
        ResolveAlarmTxtForHtml_ = webAlarmTxt
    ElseIf checkTrim(rawTxt) And Len(wendyAlarmTxt) > 0 Then
        ResolveAlarmTxtForHtml_ = wendyAlarmTxt
    Else
        ResolveAlarmTxtForHtml_ = rawTxt
    End If
End Function

Private Function checkTrim(ByVal rawTxt As String) As Boolean
    checkTrim = _
    InStr(1, rawTxt, "[截斷]", vbTextCompare) > 0 Or _
    InStr(1, rawTxt, "[截斷", vbTextCompare) > 0 Or _
    InStr(1, rawTxt, "[截", vbTextCompare) > 0
End Function

Public Function BuildRtuHtmlSubject( _
    ByVal reportDate As Date, _
    ByVal threshold As Long, _
    ByVal wsTemp As Worksheet, _
    Optional ByVal maxLen As Long = 240) As String
    Dim subject As String
    Dim dModels As Object
    Dim modelOrder As Collection
    Dim seen As Object
    Dim lastRow As Long
    Dim r As Long
    Dim eqpModel As String, pu As String, key As String
    Dim cnt As Variant
    Dim c As Collection
    Dim parts As Collection
    Dim m As Variant
    Dim p As String
    Dim subjectDate As Date
    subjectDate = reportDate
    If Not wsTemp Is Nothing Then
        Dim dt As Variant, maxDt As Date, hasDt As Boolean
        Dim r0 As Long, lastRow0 As Long
        lastRow0 = wsTemp.Cells(wsTemp.rows.Count, 4).End(xlUp).Row
        For r0 = 2 To lastRow0
            dt = wsTemp.Cells(r0, 4).Value
            If IsDate(dt) Then
                If Not hasDt Then
                    maxDt = CDate(dt)
                    hasDt = True
                ElseIf CDate(dt) > maxDt Then
                    maxDt = CDate(dt)
                End If
            End If
        Next r0
        If hasDt Then subjectDate = DateValue(maxDt)
    End If
    subject = "RTU Invalid " & Format$(subjectDate, "mm/dd") & " 機台數量 >= " & threshold & " 台"
    Set dModels = CreateObject("Scripting.Dictionary")
    Set modelOrder = New Collection
    Set seen = CreateObject("Scripting.Dictionary")
    If wsTemp Is Nothing Then
        BuildRtuHtmlSubject = subject & "(None)"
        Exit Function
    End If
    lastRow = wsTemp.Cells(wsTemp.rows.Count, 2).End(xlUp).Row
    If lastRow < 2 Then
        BuildRtuHtmlSubject = subject & "(None)"
        Exit Function
    End If
    For r = 2 To lastRow
        eqpModel = Trim$(CStr(wsTemp.Cells(r, 1).Value))
        pu = Trim$(CStr(wsTemp.Cells(r, 2).Value))
        If Len(eqpModel) = 0 Then eqpModel = "(Unknown EqpModel)"
        If Len(pu) = 0 Then GoTo NextR
        key = eqpModel & "||" & pu
        If seen.Exists(key) Then GoTo NextR
        seen.Add key, True
        cnt = wsTemp.Cells(r, 6).Value ' count colF
        If Not dModels.Exists(eqpModel) Then
            Set c = New Collection
            dModels.Add eqpModel, c
            modelOrder.Add eqpModel
        End If
        dModels(eqpModel).Add pu & " (" & CStr(cnt) & ")"
NextR:
    Next r
    If dModels.Count = 0 Then
        subject = subject & "(None)"
    Else
        Set parts = New Collection
        For Each m In modelOrder
            p = "<<" & CStr(m) & " -- " & JoinCollection_(dModels(CStr(m)), ", ") & ">>"
            parts.Add p
        Next m
        subject = subject & JoinCollection_(parts, " ")
    End If
    If maxLen > 0 And Len(subject) > maxLen Then
        subject = Left$(subject, maxLen - 3) & "..."
    End If
    BuildRtuHtmlSubject = subject
End Function

Public Function BuildRtuHtmlBody( _
    ByVal reportDate As Date, _
    ByVal threshold As Long, _
    ByVal byLotThreshold As Long, _
    ByVal wsTemp As Worksheet, _
    ByVal wsByLot As Worksheet, _
    ParamArray extraSheets() As Variant) As String

    Dim wb As Workbook
    Set wb = ResolveReportWorkbook_(wsTemp, wsByLot, extraSheets)

    Dim html As String
    html = ""
    html = html & "<!doctype html><html><head>"
    html = html & "<meta charset=""UTF-8"">"
    html = html & "<style>"
    html = html & "body{font-family:Segoe UI,Arial,sans-serif;margin:14px;color:#111;}"
    html = html & ".container{max-width:1000px;margin:0 auto;}.container.owner-wide{max-width:1200px;}" '---
    html = html & ".category{margin:0 0 22px;}"
    html = html & ".category-title{font-size:20px;font-weight:bold;margin:0 0 10px;color:red;border-bottom:1px solid #ccc;padding-bottom:4px;}"
    html = html & ".dispatch-title{color:#111;}"
    html = html & ".subcategory-title{font-size:16px;font-weight:700;color:blue;margin:12px 0 5px;}"
    html = html & ".row{font-size:13px;font-weight:400;line-height:1.35;margin:0 0 4px;white-space:pre-wrap;}"
    html = html & ".eqptype{font-size:16px;font-weight:bold;}"
    html = html & ".muted{font-size:10px;color:#666;}"
    html = html & ".toolbar{margin:0 0 18px;padding:0 0 10px;border-bottom:1px solid #ddd;}"
    html = html & ".nav-button{font:700 13px Segoe UI,Arial,sans-serif;padding:6px 12px;border:1px solid #888;background:#f7f7f7;color:#111;cursor:pointer;}"
    html = html & ".nav-button:hover{background:#eee;}"
    html = html & ".view-hidden{display:none;}"
    html = html & ".match-stats{font-size:14px;margin:0 0 16px;padding:8px 0;border-bottom:1px solid #ccc;}"
    html = html & ".match-eqptype{font-size:18px;font-weight:700;color:red;margin:18px 0 7px;}"
    html = html & ".match-item{border-bottom:1px solid #ccc;padding:5px 0;}"
    html = html & ".match-item.match-unmatched{background:#fff8cc;}"
    html = html & ".match-item summary{cursor:pointer;font-size:13px;line-height:1.45;}"
    html = html & ".match-detail{padding:4px 0 2px 20px;}"
    html = html & ".source-row{font-size:13px;line-height:1.45;white-space:pre-wrap;min-height:19px;}"
    html = html & ".source-label{display:inline-block;width:34px;font-weight:700;}"
    html = html & ".match-pu{color:blue;font-weight:700;}"
    html = html & ".diff-zero{color:green;font-weight:700;}.diff-near{color:blue;font-weight:700;}.diff-far{color:#d97706;font-weight:700;}"
    html = html & ".same-time-group{border-bottom:1px solid #999;padding:3px 0;margin:0 0 4px;}"
    html = html & ".same-time-item{padding:4px 0;}"
    html = html & ".same-time-item summary{cursor:pointer;font-size:13px;line-height:1.45;}"
    html = html & ".same-time-item.match-unmatched{background:#fff8cc;}"
    html = html & "</style>"
    html = html & "<script>"
    html = html & "function showReportView(id){"
    html = html & "var main=document.getElementById('main-view');"
    html = html & "var same=document.getElementById('same-time-view');"
    html = html & "var compare=document.getElementById('rtu-match-view');"
    html = html & "var owner=document.getElementById('owner-view');" '---
    html = html & "if(!main||!same||!compare||!owner)return;" '---
    html = html & "main.className=(id==='main')?'':'view-hidden';"
    html = html & "same.className=(id==='same-time')?'':'view-hidden';"
    html = html & "compare.className=(id==='rtu-match')?'':'view-hidden';"
    html = html & "owner.className=(id==='owner')?'':'view-hidden';" '---
    html = html & "main.parentNode.className=(id==='owner')?'container owner-wide':'container';" '---
    html = html & "if(id==='owner'){var f=document.getElementById('owner-frame');if(f&&f.contentWindow)f.contentWindow.dispatchEvent(new Event('resize'));}" '---
    html = html & "window.scrollTo(0,0);"
    html = html & "}"
    html = html & "</script></head><body><div class=""container"">"
    html = html & "<div id=""main-view"">"
    html = html & "<div class=""toolbar""><button type=""button"" class=""nav-button"" onclick=""showReportView('same-time')"">同機台同時間</button> <button type=""button"" class=""nav-button"" onclick=""showReportView('rtu-match')"">RTU比對</button> <button type=""button"" class=""nav-button"" onclick=""showReportView('owner')"">Owner</button></div>" '---
    html = html & RenderCategoryDigestHtml_(wb, "Dreams錯誤")
    html = html & RenderCategoryDigestHtml_(wb, "JI Precheck Fail")
    html = html & RenderCategoryDigestHtml_(wb, "JobinCheckFail")
    html = html & RenderCategoryDigestHtml_(wb, "RORZE")
    html = html & RenderCategoryDigestHtml_(wb, "人員介入")
    html = html & RenderCategoryDigestHtml_(wb, "IPRFail")
    html = html & RenderCategoryDigestHtml_(wb, "ReadTagFail")
    html = html & RenderCategoryDigestHtml_(wb, "No PPID")
    html = html & RenderCategoryDigestHtml_(wb, "Load Fail")
    html = html & RenderCategoryDigestHtml_(wb, "JI Fail")
    html = html & RenderCategoryDigestHtml_(wb, "其它")
    html = html & "</div>"
    html = html & "<div id=""same-time-view"" class=""view-hidden"">"
    html = html & "<div class=""toolbar""><button type=""button"" class=""nav-button"" onclick=""showReportView('main')"">返回主報表</button></div>"
    html = html & RenderSamePuSameTimeHtml_(wb)
    html = html & "</div>"
    html = html & "<div id=""rtu-match-view"" class=""view-hidden"">"
    html = html & "<div class=""toolbar""><button type=""button"" class=""nav-button"" onclick=""showReportView('main')"">返回主報表</button></div>"
    html = html & RenderRtuMatchHtml_(wb)
    html = html & "</div>"
    html = html & "<div id=""owner-view"" class=""view-hidden"">" '---
    html = html & "<div class=""toolbar""><button type=""button"" class=""nav-button"" onclick=""showReportView('main')"">返回主報表</button></div>" '---
    html = html & RenderOwner(wb) '---
    html = html & "</div>" '---
    html = html & "</div></body></html>"

    BuildRtuHtmlBody = html
    Dim desktop As String
    desktop = CreateObject("WScript.Shell").SpecialFolders("Desktop")

    Dim filePath As String
    filePath = desktop & "\RTU當日彙總_" & Format$(reportDate, "mm_dd") & ".html"

    SaveTextUtf8 filePath, html
End Function

Private Function ResolveReportWorkbook_(ByVal wsTemp As Worksheet, ByVal wsByLot As Worksheet, ByVal extraSheets As Variant) As Workbook
    Dim i As Long
    Dim candidate As Object
    If Not wsTemp Is Nothing Then
        Set ResolveReportWorkbook_ = wsTemp.Parent
        Exit Function
    End If
    If Not wsByLot Is Nothing Then
        Set ResolveReportWorkbook_ = wsByLot.Parent
        Exit Function
    End If
    On Error GoTo NoExtraSheets
    For i = LBound(extraSheets) To UBound(extraSheets)
        If IsObject(extraSheets(i)) Then
            Set candidate = extraSheets(i)
            If Not candidate Is Nothing Then
                Set ResolveReportWorkbook_ = candidate.Parent
                Exit Function
            End If
        End If
    Next i
NoExtraSheets:
End Function

Private Function GetWorksheetByName_PT(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    If wb Is Nothing Then Exit Function
    On Error Resume Next
    Set GetWorksheetByName_PT = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function RenderCategoryDigestHtml_(ByVal wb As Workbook, ByVal sheetName As String) As String
    Dim ws As Worksheet
    Set ws = GetWorksheetByName_PT(wb, sheetName)

    Dim html As String

    If ws Is Nothing Then
        RenderCategoryDigestHtml_ = html & "<div class=""muted"">(sheet not found)</div>"
        Exit Function
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.Count, 2).End(xlUp).Row
    If lastRow < 2 Then
        RenderCategoryDigestHtml_ = "<div class=""category""><div class=""category-title"">" & _
                                    HtmlEscape_(sheetName) & " (0)</div></div>"
        Exit Function
    End If

    Dim totalCount As Long: totalCount = lastRow - 1

    html = "<div class=""category""><div class=""category-title"">" & HtmlEscape_(sheetName) & " (" & totalCount & ")" & "</div>"

    Dim subOrder As Collection: Set subOrder = New Collection
    Dim rowsBySub As Object: Set rowsBySub = CreateObject("Scripting.Dictionary")
    rowsBySub.CompareMode = vbTextCompare

    Dim r As Long
    For r = 2 To lastRow
        Dim subCategory As String
        subCategory = Trim$(CStr(ws.Cells(r, 13).Value))
        If Len(subCategory) = 0 Then subCategory = "others"

        If Not rowsBySub.Exists(subCategory) Then
            Dim rowList As Collection
            Set rowList = New Collection
            rowsBySub.Add subCategory, rowList
            subOrder.Add subCategory
        End If
        rowsBySub(subCategory).Add r
    Next r

    Dim subItem As Variant
    For Each subItem In subOrder
        html = html & RenderCategorySubCategoryHtml_(ws, CStr(subItem), rowsBySub(CStr(subItem)))
    Next subItem

    RenderCategoryDigestHtml_ = html & "</div>"
End Function

Private Function RenderSamePuSameTimeHtml_(ByVal wb As Workbook) As String
    Const SHEET_CLEANED As String = "wsCleaned"
    Const REPORT_NAME As String = "同機台同時間"

    Dim ws As Worksheet
    Set ws = GetWorksheetByName_PT(wb, SHEET_CLEANED)

    If ws Is Nothing Then
        RenderSamePuSameTimeHtml_ = "<div class=""muted"">(wsCleaned sheet not found)</div>"
        Exit Function
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.Count, 4).End(xlUp).Row
    If lastRow < 2 Then
        RenderSamePuSameTimeHtml_ = "<div class=""category""><div class=""category-title"">" & _
                                    HtmlEscape_(REPORT_NAME) & " (0)</div></div>"
        Exit Function
    End If

    Dim rowsByKey As Object, keyOrder As Collection
    Set rowsByKey = CreateObject("Scripting.Dictionary")
    rowsByKey.CompareMode = vbTextCompare
    Set keyOrder = New Collection

    Dim r As Long
    For r = 2 To lastRow
        Dim pu As String, eqptype As String, timeKey As String, key As String

        pu = Trim$(CStr(ws.Cells(r, 4).Value))
        eqptype = Trim$(CStr(ws.Cells(r, 2).Value))
        If Len(eqptype) = 0 Then eqptype = "(Unknown EqpType)"
        timeKey = NormalizeTimeKey_(ws.Cells(r, 6).Value)
        If Len(pu) = 0 Or Len(timeKey) = 0 Then GoTo NextCountRow

        key = eqptype & "||" & pu & "||" & timeKey
        If Not rowsByKey.Exists(key) Then
            Dim rowList As Collection
            Set rowList = New Collection
            rowsByKey.Add key, rowList
            keyOrder.Add key
        End If
        rowsByKey(key).Add r
NextCountRow:
    Next r

    Dim duplicateRows As Long
    Dim item As Variant
    For Each item In keyOrder
        If rowsByKey(CStr(item)).Count > 1 Then duplicateRows = duplicateRows + rowsByKey(CStr(item)).Count
    Next item

    If duplicateRows = 0 Then
        RenderSamePuSameTimeHtml_ = "<div class=""category""><div class=""category-title"">" & _
                                    HtmlEscape_(REPORT_NAME) & " (0)</div></div>"
        Exit Function
    End If

    Dim eqpGroups As Object: Set eqpGroups = CreateObject("Scripting.Dictionary")
    Dim eqpOrder As Collection: Set eqpOrder = New Collection
    eqpGroups.CompareMode = vbTextCompare

    Dim parts() As String, groupKeys As Collection
    For Each item In keyOrder
        Set rowList = rowsByKey(CStr(item))
        If rowList.Count > 1 Then
            parts = Split(CStr(item), "||")
            eqptype = parts(0)
            If Not eqpGroups.Exists(eqptype) Then
                Set groupKeys = New Collection
                eqpGroups.Add eqptype, groupKeys
                eqpOrder.Add eqptype
            End If
            eqpGroups(eqptype).Add CStr(item)
        End If
    Next item

    Dim html As String
    html = "<div class=""category""><div class=""category-title"">" & _
           HtmlEscape_(REPORT_NAME) & " (" & duplicateRows & ")</div>"

    Dim eqpItem As Variant, groupKey As Variant
    For Each eqpItem In eqpOrder
        html = html & "<div class=""match-eqptype"">" & HtmlEscape_(CStr(eqpItem)) & "</div>"
        Set groupKeys = eqpGroups(CStr(eqpItem))
        For Each groupKey In groupKeys
            Set rowList = rowsByKey(CStr(groupKey))
            html = html & "<div class=""same-time-group"">"

            Dim i As Long
            For i = 1 To rowList.Count
                r = CLng(rowList(i))

                Dim alarmTxt As String, webAlarmTxt As String, wendyAlarmTxt As String
                Dim itemClass As String
                pu = Trim$(CStr(ws.Cells(r, 4).Value))
                alarmTxt = CStr(ws.Cells(r, 8).Value)
                webAlarmTxt = SanitizeOneLine_PT(CStr(ws.Cells(r, 9).Value))
                wendyAlarmTxt = SanitizeOneLine_PT(CStr(ws.Cells(r, 15).Value))
                alarmTxt = SanitizeOneLine_PT(alarmTxt)

                itemClass = "same-time-item"
                If Len(webAlarmTxt) = 0 And Len(wendyAlarmTxt) = 0 Then itemClass = itemClass & """"

                html = html & "<details class=""" & itemClass & """><summary><span class=""source-label"">RTU</span>" & _
                              "<span class=""match-pu"">" & HtmlEscape_(pu) & "</span> | " & _
                              HtmlEscape_(FormatMatchTime_(ws.Cells(r, 6).Value)) & " | " & HtmlEscape_(alarmTxt) & "</summary>"
                html = html & "<div class=""match-detail"">"
                If Len(webAlarmTxt) > 0 Then
                    html = html & BuildMatchedSourceRow_("Web", pu, ws.Cells(r, 16).Value, ws.Cells(r, 17).Value, webAlarmTxt)
                Else
                    html = html & "<div class=""source-row"">&nbsp;</div>"
                End If
                If Len(wendyAlarmTxt) > 0 Then
                    html = html & BuildMatchedSourceRow_("Wdy", pu, ws.Cells(r, 18).Value, ws.Cells(r, 19).Value, wendyAlarmTxt)
                Else
                    html = html & "<div class=""source-row"">&nbsp;</div>"
                End If
                html = html & "</div></details>"
            Next i

            html = html & "</div>"
        Next groupKey
    Next eqpItem

    RenderSamePuSameTimeHtml_ = html & "</div>"
End Function

Private Function RenderRtuMatchHtml_(ByVal wb As Workbook) As String
    Dim ws As Worksheet
    Set ws = GetWorksheetByName_PT(wb, "wsCleaned")
    If ws Is Nothing Then
    RenderRtuMatchHtml_ = "<div class=""muted"">(wsCleaned sheet not found)"
    Exit Function
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row

    Dim eqpOrder As Collection: Set eqpOrder = New Collection
    Dim puMaps As Object: Set puMaps = CreateObject("Scripting.Dictionary")
    Dim puOrders As Object: Set puOrders = CreateObject("Scripting.Dictionary")
    puMaps.CompareMode = vbTextCompare
    puOrders.CompareMode = vbTextCompare
    Dim webcol As Variant, wdycol As Variant, eqpcol As Variant, pucol As Variant
    Dim weblog As String, wdylog As String, idx As Long

    eqpcol = ws.Range("B2:B" & lastRow).Value2
    pucol = ws.Range("D2:D" & lastRow).Value2
    webcol = ws.Range("I2:I" & lastRow).Value2
    wdycol = ws.Range("O2:O" & lastRow).Value2

    Dim totalCount As Long, webCount As Long, wendyCount As Long, noMatch As Long
    Dim r As Long
    For r = 2 To lastRow
        idx = r - 1
        Dim eqptype As String, pu As String
        eqptype = Trim$(CStr(eqpcol(idx, 1)))
        pu = Trim$(CStr(pucol(idx, 1)))
        If Len(eqptype) = 0 Then eqptype = "(Unknown EqpType)"
        If Len(pu) = 0 Then GoTo NextMatchRow

        totalCount = totalCount + 1
        
        weblog = webcol(idx, 1)
        wdylog = wdycol(idx, 1)
        
        
        If Len(weblog) > 0 Then webCount = webCount + 1
        If Len(wdylog) > 0 Then wendyCount = wendyCount + 1
        If Len(weblog) < 1 And Len(wdylog) < 1 Then noMatch = noMatch + 1
        
        Dim puMap As Object, puOrder As Collection, rowList As Collection
        If Not puMaps.Exists(eqptype) Then
            Set puMap = CreateObject("Scripting.Dictionary")
            puMap.CompareMode = vbTextCompare
            Set puOrder = New Collection
            puMaps.Add eqptype, puMap
            puOrders.Add eqptype, puOrder
            eqpOrder.Add eqptype
        Else
            Set puMap = puMaps(eqptype)
            Set puOrder = puOrders(eqptype)
        End If

        If Not puMap.Exists(pu) Then
            Set rowList = New Collection
            puMap.Add pu, rowList
            puOrder.Add pu
        Else
            Set rowList = puMap(pu)
        End If
        rowList.Add r
NextMatchRow:
    Next r

    Dim html As String
    html = "<div class=""match-stats"">RTU總數：" & totalCount & _
        "　WebCSV總數：" & webCount & "　Wendy總數： " & wendyCount & _
        " &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp  未命中數  ： " & noMatch & "</div>"

    Dim eqpItem As Variant, puItem As Variant
    For Each eqpItem In eqpOrder
        html = html & "<div class=""match-eqptype"">" & HtmlEscape_(CStr(eqpItem)) & "</div>"
        Set puMap = puMaps(CStr(eqpItem))
        Set puOrder = puOrders(CStr(eqpItem))
        For Each puItem In puOrder
            html = html & RenderRtuMatchRows_(ws, CStr(puItem), puMap(CStr(puItem)))
        Next puItem
    Next eqpItem

    If totalCount = 0 Then html = html & "<div class=""muted"">(none)</div>"
    RenderRtuMatchHtml_ = html
End Function

Private Function RenderRtuMatchRows_(ByVal ws As Worksheet, ByVal pu As String, ByVal rowList As Collection) As String
    Dim indexes() As Long
    ReDim indexes(1 To rowList.Count)
    Dim i As Long, j As Long, temp As Long
    For i = 1 To rowList.Count: indexes(i) = CLng(rowList(i)): Next i

    For i = 2 To UBound(indexes)
        j = i
        Do While j > 1
        If MatchRowComesLater_(ws, indexes(j - 1), indexes(j)) Then
            temp = indexes(j - 1): indexes(j - 1) = indexes(j): indexes(j) = temp
            j = j - 1
        Else
            Exit Do
        End If
        
        Loop
    Next i

    Dim html As String
    For i = 1 To UBound(indexes)
        Dim r As Long: r = indexes(i)
        Dim rtuText As String, webText As String, wendyText As String
        rtuText = SanitizeOneLine_PT(CStr(ws.Cells(r, 8).Value))
        webText = SanitizeOneLine_PT(CStr(ws.Cells(r, 9).Value))
        wendyText = SanitizeOneLine_PT(CStr(ws.Cells(r, 15).Value))

        Dim matchItemClass As String
        matchItemClass = "match-item"
        If Len(webText) = 0 And Len(wendyText) = 0 Then
            matchItemClass = matchItemClass & " match-unmatched"
        End If

        html = html & "<details class=""" & matchItemClass & """><summary><span class=""source-label"">RTU</span>" & _
                    "<span class=""match-pu"">" & HtmlEscape_(pu) & "</span> | " & _
                    HtmlEscape_(FormatMatchTime_(ws.Cells(r, 6).Value)) & " | " & HtmlEscape_(rtuText) & "</summary>"
        html = html & "<div class=""match-detail"">"

        If Len(webText) > 0 Then
            html = html & BuildMatchedSourceRow_("Web", pu, ws.Cells(r, 16).Value, ws.Cells(r, 17).Value, webText)
        Else
            html = html & "<div class=""source-row"">&nbsp;</div>"
        End If

        If Len(wendyText) > 0 Then
            html = html & BuildMatchedSourceRow_("Wdy", pu, ws.Cells(r, 18).Value, ws.Cells(r, 19).Value, wendyText)
        Else
            html = html & "<div class=""source-row"">&nbsp;</div>"
        End If

        html = html & "</div></details>"
    Next i
    RenderRtuMatchRows_ = html
End Function

Private Function BuildMatchedSourceRow_(ByVal sourceLabel As String, ByVal pu As String, ByVal matchedTime As Variant, ByVal diffSeconds As Variant, ByVal alarmtext As String) As String
    BuildMatchedSourceRow_ = "<div class=""source-row""><span class=""source-label"">" & _
                             HtmlEscape_(sourceLabel) & "</span><span class=""match-pu"">" & _
                             HtmlEscape_(pu) & "</span> | " & HtmlEscape_(FormatMatchTime_(matchedTime)) & _
                             FormatDiffSecondsHtml_(diffSeconds) & " | " & HtmlEscape_(alarmtext) & "</div>"
End Function

Private Function FormatDiffSecondsHtml_(ByVal v As Variant) As String
    If Not IsNumeric(v) Then Exit Function
    Dim seconds As Long: seconds = CLng(v)
    Dim cssClass As String, displayText As String
    If seconds = 0 Then
        cssClass = "diff-zero": displayText = "0秒"
    ElseIf Abs(seconds) <= 30 Then
        cssClass = "diff-near"
    Else
        cssClass = "diff-far"
    End If
    If Len(displayText) = 0 Then
        If seconds > 0 Then displayText = "+" & seconds & "秒" Else displayText = CStr(seconds) & "秒"
    End If
    FormatDiffSecondsHtml_ = "（<span class=""" & cssClass & """>" & displayText & "</span>）"
End Function

Private Function FormatMatchTime_(ByVal v As Variant) As String
    If IsDate(v) Then
        FormatMatchTime_ = Format$(CDate(v), "yyyy/mm/dd hh:nn:ss")
    Else
        FormatMatchTime_ = Trim$(CStr(v))
    End If
End Function

Private Function MatchRowComesLater_(ByVal ws As Worksheet, ByVal leftRow As Long, ByVal rightRow As Long) As Boolean
    Dim leftTime As Double, rightTime As Double
    If IsDate(ws.Cells(leftRow, 6).Value) Then leftTime = CDbl(CDate(ws.Cells(leftRow, 6).Value)) Else leftTime = 1E+99
    If IsDate(ws.Cells(rightRow, 6).Value) Then rightTime = CDbl(CDate(ws.Cells(rightRow, 6).Value)) Else rightTime = 1E+99
    If leftTime <> rightTime Then MatchRowComesLater_ = (leftTime > rightTime) Else MatchRowComesLater_ = (leftRow > rightRow)
End Function

Private Function NormalizeTimeKey_(ByVal v As Variant) As String
    If IsDate(v) Then
        NormalizeTimeKey_ = Format$(CDate(v), "yyyy-mm-dd hh:nn:ss")
    Else
        NormalizeTimeKey_ = Trim$(CStr(v))
    End If
End Function

Private Function FormatTimeForHtml_(ByVal v As Variant) As String
    If IsDate(v) Then
        FormatTimeForHtml_ = Format$(CDate(v), "hh:nn")
    Else
        FormatTimeForHtml_ = Trim$(CStr(v))
    End If
End Function

Private Function RenderCategorySubCategoryHtml_(ByVal ws As Worksheet, ByVal subCategory As String, ByVal rowList As Collection) As String
    Dim html As String
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    html = "<div class=""subcategory-title"">" & HtmlEscape_(subCategory) & " (" & rowList.Count & ")" & "</div>"

    Dim i As Long
    For i = 1 To rowList.Count
        Dim r As Long
        Dim eqptype As String
        Dim pu As String
        Dim alarmTxt As String
        Dim fullAlarmTxt As String
        Dim wendyAlarmTxt As String
        Dim resolvedAlarm As String
        Dim key As String
        Dim alarmTmst As String
        
        r = CLng(rowList(i))
        eqptype = Trim$(CStr(ws.Cells(r, 1).Value))
        pu = Trim$(CStr(ws.Cells(r, 2).Value))
        If Len(eqptype) = 0 Then eqptype = "(Unknown EqpType)"
        If Len(pu) = 0 Then GoTo NextRow

        key = subCategory & "||" & eqptype & "||" & pu
        If seen.Exists(key) Then GoTo NextRow
        seen.Add key, True
        
        If IsDate(ws.Cells(r, 4).Value) Then
            alarmTmst = Format$(CDate(ws.Cells(r, 4).Value), "hh:nn")
        Else
            alarmTmst = ""
        End If
        alarmTxt = CStr(ws.Cells(r, 5).Value)
        fullAlarmTxt = CStr(ws.Cells(r, 8).Value)
        wendyAlarmTxt = CStr(ws.Cells(r, 14).Value)
        resolvedAlarm = SanitizeOneLine_PT(ResolveAlarmTxtForHtml_(alarmTxt, fullAlarmTxt, wendyAlarmTxt))

        html = html & "<div class=""row""><span class=""eqptype"">" & HtmlEscape_(eqptype) & "</span> " & _
                    HtmlEscape_(pu & " | " & alarmTmst & " | " & resolvedAlarm) & "</div>"
NextRow:
    Next i

    If seen.Count = 0 Then html = html & "<div class=""muted"">(none)</div>"
    RenderCategorySubCategoryHtml_ = html
End Function

' -------------------------
' HTML Helpers
' -------------------------
Private Function HtmlEscape_(ByVal s As String) As String
    s = CStr(s)
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, """", "&quot;")
    s = Replace(s, "'", "&#39;")
    HtmlEscape_ = s
End Function
Private Function JoinCollection_(ByVal c As Collection, ByVal sep As String) As String
    Dim i As Long, out As String
    out = ""
    For i = 1 To c.Count
        If i > 1 Then out = out & sep
        out = out & CStr(c(i))
    Next i
    JoinCollection_ = out
End Function

Public Sub SaveTextUtf8(ByVal filePath As String, ByVal text As String)
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")

    stm.Type = 2           ' adTypeText
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText text
    stm.SaveToFile filePath, 2   ' adSaveCreateOverWrite
    stm.Close
End Sub




' Owner dashboard: all data is embedded; no workbook access from the browser.
'----------------
Private Function RenderOwner(ByVal wb As Workbook) As String
    Dim page As String
    page = "<!doctype html><html lang=""zh-Hant""><head><meta charset=""utf-8""><meta name=""viewport"" content=""width=device-width,initial-scale=1""><style>"
    page = page & OwnerStyles_() & "</style></head><body>" & OwnerMarkup_()
    page = page & "<script>const data=" & OwnerDataJson_(wb) & ";" & OwnerScript_() & "</script></body></html>"
    RenderOwner = "<iframe id=""owner-frame"" title=""Owner 七天報表"" style=""display:block;width:100%;height:100px;border:0;"" srcdoc=""" & HtmlEscape_(page) & """></iframe>"
End Function

'----------------
Private Function OwnerString_(ByVal value As Variant) As String
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then Exit Function
    OwnerString_ = Trim$(CStr(value))
End Function

'----------------
Private Function OwnerJsonString_(ByVal value As String) As String
    Dim i As Long, code As Long, ch As String, out As String
    For i = 1 To Len(value)
        ch = Mid$(value, i, 1)
        code = AscW(ch)
        Select Case code
            Case 34: out = out & "\"""
            Case 92: out = out & "\\"
            Case 0 To 31, 60, 38, &H2028, &H2029
                out = out & "\u" & Right$("0000" & Hex$(code), 4)
            Case Else: out = out & ch
        End Select
    Next i
    OwnerJsonString_ = """" & out & """"
End Function

'----------------
Private Function OwnerDataJson_(ByVal wb As Workbook) As String
    Dim ws As Worksheet, stats As Worksheet, owners As Object
    Set ws = GetWorksheetByName_PT(wb, "accumulated")
    Set stats = GetWorksheetByName_PT(wb, "7-day")
    Set owners = buildpuowner()
    Dim days As Object, counts As Object, types As Object
    Set days = CreateObject("Scripting.Dictionary")
    Set counts = CreateObject("Scripting.Dictionary")
    Set types = CreateObject("Scripting.Dictionary")
    counts.CompareMode = vbTextCompare
    types.CompareMode = vbTextCompare
    Dim errorText As String, lastCol As Long, firstCol As Long, c As Long, r As Long
    Dim dt As Date, baseDate As Date, key As String, pu As String, v As Variant
    baseDate = Date
    If ws Is Nothing Then errorText = "找不到 accumulated 工作表。"
    If stats Is Nothing Then
        If Len(errorText) > 0 Then errorText = errorText & " "
        errorText = errorText & "找不到 7-day 工作表。"
    Else
        lastCol = stats.Cells(1, stats.Columns.Count).End(xlToLeft).Column
        firstCol = lastCol - 6
        If firstCol < 2 Then firstCol = 2
        For c = firstCol To lastCol
            v = stats.Cells(1, c).Value
            If Not IsError(v) Then
                If IsDate(v) Then
                    dt = DateValue(CDate(v))
                    If dt >= baseDate - 7 And dt <= baseDate - 1 Then
                        key = Format$(dt, "yyyy/mm/dd")
                        If Not days.Exists(key) Then days.Add key, c
                    End If
                End If
            End If
        Next c
        If days.Count = 0 Then
            If Len(errorText) > 0 Then errorText = errorText & " "
            errorText = errorText & "7-day 最後七欄沒有 Date-7 至 Date-1 的有效日期。"
        End If
    End If
    Dim dateKeys As Variant, swap As Variant, i As Long, j As Long
    dateKeys = days.Keys
    For i = 0 To days.Count - 2
        For j = i + 1 To days.Count - 1
            If CStr(dateKeys(i)) > CStr(dateKeys(j)) Then
                swap = dateKeys(i): dateKeys(i) = dateKeys(j): dateKeys(j) = swap
            End If
        Next j
    Next i
    Dim datesJson As String, countKey As String, n As Double
    datesJson = "["
    For i = 0 To days.Count - 1
        If i > 0 Then datesJson = datesJson & ","
        datesJson = datesJson & OwnerJsonString_(CStr(dateKeys(i)))
    Next i
    datesJson = datesJson & "]"
    If Not stats Is Nothing Then
        For r = 2 To 360
            pu = OwnerString_(stats.Cells(r, 1).Value)
            If Len(pu) > 0 Then
                For i = 0 To days.Count - 1
                    key = CStr(dateKeys(i))
                    countKey = pu & "|" & key
                    v = stats.Cells(r, CLng(days(key))).Value
                    n = 0
                    If Not IsError(v) Then
                        If IsNumeric(v) Then
                            n = CDbl(v)
                            If n < 0 Then n = 0
                        End If
                    End If
                    counts(countKey) = n
                Next i
            End If
        Next r
    End If
    Dim logParts As Collection: Set logParts = New Collection
    Dim lastRow As Long, category As String, subCategory As String, alarm As String, eqptype As String
    If Not ws Is Nothing Then
        lastRow = ws.Cells(ws.rows.Count, 4).End(xlUp).Row
        Dim values As Variant
        If lastRow >= 2 Then
            values = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 23)).Value
            For r = 1 To UBound(values, 1)
                pu = OwnerString_(values(r, 4))
                If owners.Exists(pu) Then
                    eqptype = OwnerString_(values(r, 2))
                    If Len(eqptype) > 0 And Not types.Exists(pu) Then types.Add pu, eqptype
                    v = values(r, 6)
                    If Not IsError(v) Then
                        If IsDate(v) Then
                            dt = DateValue(CDate(v))
                            key = Format$(dt, "yyyy/mm/dd")
                            If days.Exists(key) Then
                                category = OwnerString_(values(r, 11))
                                subCategory = OwnerString_(values(r, 22))
                                If Len(category) = 0 Then category = "其它"
                                If Len(subCategory) = 0 Then subCategory = "others"
                                alarm = SanitizeOneLine_PT(ResolveAlarmTxtForHtml_(OwnerString_(values(r, 8)), OwnerString_(values(r, 9)), OwnerString_(values(r, 15))))
                                logParts.Add "[" & OwnerJsonString_(LCase$(pu)) & "," & OwnerJsonString_(key) & "," & OwnerJsonString_(category) & "," & OwnerJsonString_(subCategory) & "," & OwnerJsonString_(Format$(CDate(v), "yyyy/mm/dd hh:nn:ss")) & "," & OwnerJsonString_(alarm) & "]"
                            End If
                        End If
                    End If
                End If
            Next r
        End If
    End If
    Dim machineParts As Collection: Set machineParts = New Collection
    Dim p As Variant, series As String
    For Each p In owners.Keys
        pu = CStr(p)
        eqptype = "未分類"
        If types.Exists(pu) Then eqptype = CStr(types(pu))
        series = "["
        For i = 0 To days.Count - 1
            If i > 0 Then series = series & ","
            countKey = pu & "|" & CStr(dateKeys(i))
            n = 0
            If counts.Exists(countKey) Then n = CDbl(counts(countKey))
            series = series & Trim$(Str$(n))
        Next i
        series = series & "]"
        machineParts.Add "[" & OwnerJsonString_(pu) & "," & OwnerJsonString_(eqptype) & "," & series & "," & OwnerJsonString_(CStr(owners(pu))) & "]"
    Next p
    OwnerDataJson_ = "{""dates"":" & datesJson & ",""machines"":[" & JoinCollection_(machineParts, ",") & "],""logs"":[" & JoinCollection_(logParts, ",") & "],""error"":" & OwnerJsonString_(errorText) & "}"
End Function


'----------------
Private Sub AddOwnerPus_(ByVal owners As Object, ByVal ownerName As String, ByVal list As String)
    Dim pu As Variant
    For Each pu In Split(list, ",")
        If Len(CStr(pu)) > 0 And Not owners.Exists(CStr(pu)) Then owners.Add CStr(pu), ownerName
    Next pu
End Sub

'----------------
Private Function OwnerStyles_() As String
    Dim html As String
    html = html & ":root { --ink: #23334a; --muted: #8290a2; --line: #e6ebf1; --blue: #346ee8 } * { box-sizing: border-box } body { margin: 0; background: #f4f6fa; color: var(--ink); font: 14px 'Segoe UI', 'Microsoft JhengHei', sans-serif } button, input { font: inherit } button { cursor: pointer } header { height: 78px; background: white; border-bottom: 1px solid va"
    html = html & "r(--line); display: flex; align-items: center; justify-content: space-between; padding: 0 28px } .brand { display: flex; align-items: center; gap: 13px } .logo { background: var(--blue); color: white; border-radius: 12px; padding: 10px; font-size: 24px } h1 { font-size: 19px; margin: 0 0 5px } small, .muted { color: var(--muted) } .demo { backgroun"
    html = html & "d: #edf3ff; color: #4774cb; padding: 6px 11px; border-radius: 6px; margin-right: 18px } main { height: calc(100dvh - 78px); padding: 20px; display: flex; min-height: 550px } .left { width: 27%; min-width: 220px; display: flex; flex-direction: column; gap: 16px } .panel { background: white; border: 1px solid var(--line); border-radius: 12px; overflo"
    html = html & "w: hidden; box-shadow: 0 3px 12px #23334a03 } .machine-panel { display: flex; flex-direction: column; min-height: 0; flex: 3 } .machine-panel.quiet { flex: 2 } .section-head { padding: 19px 18px 13px; display: flex; justify-content: space-between; align-items: center; font-weight: 600 } .count { font-size: 12px; background: #f0f3f8; padding: 4px 8p"
    html = html & "x; border-radius: 5px } .orange { color: #d77835 } .gray{ color:gray } .list { overflow: auto; padding: 0 9px 9px } .machine { width: 100%; background: white; border: 1px solid transparent; padding: 12px 11px; margin: 3px 0; border-radius: 8px; text-align: left; display: flex; align-items: center; justify-content: space-between } .machine:hover { b"
    html = html & "ackground: #f5f8ff } .machine.active { background: #edf3ff; border-color: #b8ceff } .machine-name { font-weight: 650; letter-spacing: .4px } .machine small { display: block; font-size: 11px; margin-top: 5px } .badge { color: #d36a30; background: #fff0e6; border-radius: 6px; min-width: 35px; text-align: center; padding: 5px } .quiet .machine { color"
    html = html & ": #919bab } .quiet .machine.active { color: var(--blue) } .quiet .badge { background: #f2f4f7; color: #9aa3b1 } .split { width: 17px; flex-shrink: 0; cursor: col-resize; display: flex; align-items: center; justify-content: center; touch-action: none } .split:after { content: ''; height: 44px; width: 4px; border-radius: 4px; background: #c8d1df } .s"
    html = html & "plit:hover:after, .dragging .split:after { background: var(--blue) } .dragging { cursor: col-resize; user-select: none } .right { min-width: 0; flex: 1; display: grid; grid-template-rows: 1fr 2fr; gap: 16px } .chart-panel { display: flex; flex-direction: column; min-height: 0 } .chart-head { display: flex; justify-content: space-between; align-item"
    html = html & "s: center; padding: 18px 23px 0 } .title { font-weight: 650; font-size: 16px } .chart-head small { display: block; margin-top: 5px } .legend { font-size: 12px; color: var(--muted) } .legend:before { content: '●'; color: var(--blue); margin-right: 7px } .chart { flex: 1; min-height: 85px; padding: 4px 20px 0 } svg { height: 100%; width: 100%; overfl"
    html = html & "ow: visible } .date-controls { display: flex; justify-content: center; align-items: center; gap: 15px; padding: 9px; border-top: 1px solid #f1f4f8; font-size: 12px } .date-controls button { border: 1px solid var(--line); border-radius: 6px; background: white; padding: 5px 10px; color: #536780 } .date-controls button:disabled { opacity: .3; cursor: "
    html = html & "default } .date-controls strong { min-width: 92px; text-align: center } .log-panel { display: flex; flex-direction: column; min-height: 0 } .log-head { display: flex; align-items: center; justify-content: space-between; gap: 12px; padding: 18px 21px; border-bottom: 1px solid var(--line) } .log-head small { display: block; margin-top: 6px } .tools {"
    html = html & " display: flex; align-items: center; gap: 10px; font-size: 12px; color: #738198 } .tools input { accent-color: var(--blue) } .log-scroll { flex: 1; overflow: auto; min-width: 0 } .log-table { border-collapse: collapse; width: 100%; font: 12px 'Consolas', 'Microsoft JhengHei', monospace } .log-table th { background: #f7f9fc; position: sticky; top: 0"
    html = html & "; color: #8290a2; text-align: left; font-size: 11px; font-weight: 500; padding: 10px 17px; border-bottom: 1px solid var(--line) } .log-table td { padding: 13px 17px; border-bottom: 1px solid #edf0f5; white-space: nowrap; vertical-align: top } .log-table tr:hover td { background: #f8faff } .log-table td:last-child { width: 100%; white-space: pre } ."
    html = html & "wrap td:last-child { white-space: pre-wrap; overflow-wrap: anywhere; min-width: 240px } .level { color: #d87945; background: #fff1e8; padding: 3px 6px; border-radius: 4px; font-size: 10px } .empty { text-align: center; padding: 65px; color: #8d99aa } .footer { padding: 10px 20px; border-top: 1px solid var(--line); font-size: 11px; color: #94a0b1; d"
    html = html & "isplay: flex; justify-content: space-between } @media(max-width:700px) { header { padding: 0 14px } .demo { display: none } main { padding: 10px; overflow: auto } .left { min-width: 170px; width: 32% } .section-head { padding: 12px; font-size: 12px } .machine { padding: 9px 5px } .machine small { font-size: 10px } .right { min-width: 350px } .log-h"
    html = html & "ead { padding: 12px } .tools { gap: 3px } .split { width: 12px } } body { background: #fff } .panel { border: 0; border-radius: 0; box-shadow: none; background: transparent } main { gap: 22px } .left { gap: 0; width: 27%; flex-shrink: 0; border-right: 1px solid #dce2eb; padding-right: 20px } .quiet { border-top: 1px solid #dce2eb; margin-top: 12px;"
    html = html & " padding-top: 6px } .right { gap: 0 } .chart-panel { border-bottom: 1px solid #dce2eb } .machine { border-radius: 0; margin: 0; border: 0; border-left: 3px solid transparent } .machine.active { border: 0; border-left: 3px solid #346ee8; background: #edf3ff } .badge, .count, .demo { border-radius: 0; background: transparent } .logo { border-radius: "
    html = html & "0 } .log-head { border-bottom: 0; padding: 17px 22px 12px } .date-controls button { border-radius: 0 } .log-scroll { overflow-y: auto; overflow-x: hidden; padding: 0 22px 20px } .log-group { margin: 0 0 15px } .group-title { color: #c42e35; font-weight: 750; font-size: 18px; margin: 7px 0 7px; padding-bottom: 8px; border-bottom: 1px solid #e8ebef }"
    html = html & " .group-title small { font-size: 12px; color: #a7777b; font-weight: 400; margin-left: 12px } .sub-title { color: #2466b5; font-size: 14px; font-weight: 600; margin: 10px 0 6px } .sub-title small { color: #7393b6; font-weight: 400; margin-left: 8px } .log-line { display: grid; grid-template-columns: 105px minmax(0, 1fr); gap: 14px; padding: 4px 0; b"
    html = html & "order-bottom: 1px solid #f0f2f5; font: 12px/1.7 Consolas, 'Microsoft JhengHei', monospace } .log-time { color: #8a96a5; white-space: nowrap } .log-text { white-space: pre-wrap; overflow-wrap: anywhere; color: #455064; min-width: 0 } .footer { border-radius: 0 } .section-head { padding-left: 10px } .chart-head { padding-top: 12px } @media(max-width:"
    html = html & "700px) { main { gap: 12px; min-width: 650px } .left { min-width: 180px; padding-right: 10px } .right { min-width: 0 } .log-line { grid-template-columns: 90px minmax(0, 1fr); gap: 8px } } html { overflow-y: scroll; background: #f7f6f2 } body { background: #f7f6f2 } header { background: #f7f6f2 } main { height: auto; min-height: calc(100dvh - 78px); "
    html = html & "align-items: flex-start; overflow: visible } .left { display: block } .machine-panel, .machine-panel.quiet { display: block; flex: none; overflow: visible } .list { overflow: visible } .machine { background: transparent } .right { display: flex; flex-direction: column; min-width: 0 } .chart-panel { height: calc((100dvh - 118px)/3); min-height: 220p"
    html = html & "x; flex: none } .log-panel { display: block; overflow: visible; min-height: calc((100dvh - 118px)*2/3) } .log-scroll { overflow: visible } .panel { overflow: visible } .quiet { margin-top: 18px } .log-line { border-bottom-color: #e8e8e2 } .group-title { border-bottom-color: #e1e2dd } .footer { margin-top: 12px } .date-controls { border-top-color: #"
    html = html & "e8e8e2 } .date-controls button { background: #fbfaf7 } .machine:hover { background: #eeeee9 } .machine.active { background: #e9eef8 } @media(max-width:700px) { main { overflow: visible } .chart-panel { min-height: 220px } } .left { width: 22%; min-width: 210px } .chart-panel { height: clamp(290px, 44vh, 440px); min-height: 290px } .log-panel { min-"
    html = html & "height: calc((100dvh - 118px)*.56) } @media(max-width:700px) { .left { width: 22%; min-width: 170px } .chart-panel { height: 290px; min-height: 290px } } .svg-grid-line { stroke: #dedfdc; stroke-width: 1 } .svg-axis { stroke: #aeb6bf; stroke-width: 1 } .svg-label { fill: #7a8797; font-size: 13px } .svg-value-label { fill: #526780; font-size: 13px; "
    html = html & "font-weight: 500 } .svg-line { fill: none; stroke-width: 2; stroke-linecap: round; stroke-linejoin: round } .svg-point { stroke: none } .chart { padding: 8px 20px 0 } .chart svg { display: block } .chart-panel { height: auto; min-height: 0 } .chart { flex: none; width: calc(100% - 40px); max-width: 700px; aspect-ratio: 8 / 5; padding: 0; margin: 10"
    html = html & "px auto 0; background: #fdfdfb } .chart svg { font-family: Calibri, 'Segoe UI', 'Microsoft JhengHei', sans-serif } .svg-grid-line { stroke: #d9dce1; stroke-width: 1 } .svg-axis { stroke: #a6adb7; stroke-width: 1 } .svg-label { fill: #66717d; font-size: 12px } .svg-value-label { fill: #3d4652; font-size: 12px; font-weight: 400 } .svg-line { stroke-w"
    html = html & "idth: 2.5 } .svg-point { stroke: #4472c4; stroke-width: 1 } #chart [data-day] { cursor: pointer; outline: none } .svg-point-halo { fill: #4472c4; opacity: 0; pointer-events: none; transition: opacity 150ms ease, fill 150ms ease } #chart .svg-point { transition: fill 150ms ease, stroke 150ms ease } #chart [data-day]:hover .svg-point-halo, #chart [da"
    html = html & "ta-day]:focus-visible .svg-point-halo { opacity: .16 } #chart [data-day].selected .svg-point-halo { fill: #ed8936; opacity: .2 } #chart [data-day].selected .svg-point { fill: #ed8936; stroke: #ed8936 } #chart [data-day].selected:hover .svg-point-halo, #chart [data-day].selected:focus-visible .svg-point-halo { opacity: .32 } main { width: 100%; max-"
    html = html & "width: 1200px; min-width: 0; min-height: 100dvh; margin: 0 auto; padding: 24px 32px; gap: 22px } .left { width: 200px; min-width: 200px; flex: 0 0 200px; padding-right: 12px } .machine-name { white-space: nowrap; letter-spacing: 0 } .machine-total { font-weight: 400 } .equipment-group { margin: 2px 0 5px } .equipment-type { margin: 0; padding: 3px "
    html = html & "11px; font-size: 12px; font-weight: 650; color: var(--muted) } .equipment-machines { margin-left: 12px; border-left: 1px solid var(--line) } .equipment-machines .machine { padding: 2px 8px } .log-panel { min-height: 0 } .log-line { display: block } @media(max-width:700px) { main { flex-direction: column } .left { width: 100%; min-width: 0; flex: no"
    html = html & "ne; border-right: 0; padding-right: 0 } .right { width: 100% } } html {overflow-y:hidden} main {min-height:0} .owner-toolbar {max-width:1200px;margin:0 auto;padding:12px 32px 0;display:flex;flex-wrap:wrap;gap:8px} .owner-button {padding:6px 12px;border:1px solid #b8c3d3;background:transparent;color:var(--ink)} .owner-button.active {background:#e9ee"
    html = html & "f8;border-color:#346ee8;color:#346ee8} .owner-extra {margin-left:24px} .owner-first,.owner-extra .log-text {cursor:text} [hidden] {display:none!important}"
    OwnerStyles_ = html
End Function

'----------------
Private Function OwnerMarkup_() As String
    Dim html As String
    html = html & "<div class=""owner-toolbar"" id=""owners""></div><div class=""empty"" id=""status"" hidden></div>"
    html = html & "<main id=""dashboard"" hidden>"
    html = html & "    <aside class=""left"" id=""left"">"
    html = html & "      <section class=""panel machine-panel"">"
    html = html & "        <div class=""section-head"">"
    html = html & "          <span><span class=""orange"">●</span> 最後一日有錯誤</span>"
    html = html & "          <span class=""count"" id=""alerts-count""></span>"
    html = html & "        </div>"
    html = html & "        <div class=""list"" id=""alerts""></div>"
    html = html & "      </section>"
    html = html & "      <section class=""panel machine-panel quiet"">"
    html = html & "        <div class=""section-head"">"
    html = html & "          <span><span class=""gray"">●</span> 最後一日無錯誤</span>"
    html = html & "          <span class=""count"" id=""quiet-count""></span>"
    html = html & "        </div>"
    html = html & "        <div class=""list"" id=""quiet""></div>"
    html = html & "      </section>"
    html = html & "    </aside>"
    html = html & "    <section class=""right"">"
    html = html & "      <section class=""panel chart-panel"">"
    html = html & "        <div class=""chart-head"">"
    html = html & "          <div>"
    html = html & "            <div class=""title"">"
    html = html & "              <span id=""chart-name""></span>"
    html = html & "              <span style=""font-weight:400;color:#b3bdca"">/</span>"
    html = html & "              近七日錯誤趨勢"
    html = html & "            </div>"
    html = html & "          </div>"
    html = html & "        </div>"
    html = html & "        <div class=""chart"">"
    html = html & "          <svg aria-label=""近七日錯誤數折線圖"" id=""chart"" preserveaspectratio=""xMidYMid meet"" viewbox=""0 0 940 470""></svg>"
    html = html & "        </div>"
    html = html & "        <div class=""date-controls"">"
    html = html & "          <button id=""prev"">← 上一日</button>"
    html = html & "          <strong id=""date""></strong>"
    html = html & "          <button id=""next"">下一日 →</button>"
    html = html & "          <button id=""today"">最後一日</button>"
    html = html & "        </div>"
    html = html & "      </section>"
    html = html & "      <section class=""panel log-panel"">"
    html = html & "        <div class=""log-head"">"
    html = html & "          <div>"
    html = html & "            <small id=""log-context""><span id=""log-context-details""></span><span class=""count"" id=""log-count""></span></small>"
    html = html & "          </div>"
    html = html & "        </div>"
    html = html & "        <div class=""log-scroll"" id=""log-scroll""></div>"
    html = html & "      </section>"
    html = html & "    </section>"
    html = html & "  </main>"
    OwnerMarkup_ = html
End Function

'----------------
Private Function OwnerScript_() As String
    Dim html As String
    html = html & "(() => {"
    html = html & "  'use strict';"
    html = html & "  const $ = id => document.getElementById(id);"
    html = html & "  const esc = value => String(value).replace(/[&<>""']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','""':'&quot;',""'"":'&#39;'}[c]));"
    html = html & "  const dates = data.dates;"
    html = html & "  let machines = [], selected = -1, day = dates.length - 1, owner = '';"
    html = html & "  const logIndex = new Map();"
    html = html & "  data.logs.forEach(row => {"
    html = html & "    const key = JSON.stringify([row[0], row[1]]);"
    html = html & "    if (!logIndex.has(key)) logIndex.set(key, []);"
    html = html & "    logIndex.get(key).push(row);"
    html = html & "  });"
    html = html & "  const ownerNames = [...new Set(data.machines.map(m => m[3]))];"
    html = html & "  $('owners').innerHTML = ownerNames.map((name,i) => `<button type=""button"" class=""owner-button"" data-owner=""${i}"">${esc(name)}</button>`).join('');"
    html = html & "  $('owners').querySelectorAll('[data-owner]').forEach(button => button.onclick = () => {"
    html = html & "    owner = ownerNames[+button.dataset.owner];"
    html = html & "    $('owners').querySelectorAll('button').forEach(b => b.classList.toggle('active', b === button));"
    html = html & "    if (data.error) return;"
    html = html & "    machines = data.machines.filter(m => m[3] === owner);"
    html = html & "    /* Match the first button actually displayed: positive-count section first. */"
    html = html & "    selected = machines.findIndex(m => m[2][dates.length - 1] > 0);"
    html = html & "    if (selected < 0) selected = 0;"
    html = html & "    day = dates.length - 1;"
    html = html & "    $('dashboard').hidden = false;"
    html = html & "    lists(); chart(); logs(); resizeFrame();"
    html = html & "  });"
    html = html & "  if (data.error) {"
    html = html & "    $('status').hidden = false;"
    html = html & "    $('status').textContent = data.error;"
    html = html & "  }"
    html = html & "  function lists() {"
    html = html & "    ['alerts','quiet'].forEach(id => {"
    html = html & "      const groups = new Map();"
    html = html & "      machines.forEach((m,i) => {"
    html = html & "        if ((m[2][dates.length - 1] > 0) !== (id === 'alerts')) return;"
    html = html & "        if (!groups.has(m[1])) groups.set(m[1], []);"
    html = html & "        groups.get(m[1]).push({m,i});"
    html = html & "      });"
    html = html & "      $(id + '-count').textContent = [...groups.values()].reduce((n,g) => n + g.length,0) + ' 台';"
    html = html & "      $(id).innerHTML = [...groups].map(([type,entries]) => `<section class=""equipment-group""><h3 class=""equipment-type"">${esc(type)}</h3><div class=""equipment-machines"">${entries.map(({m,i}) => `<button type=""button"" class=""machine ${i===selected?'active':''}"" data-machine=""${i}""><span class=""machine-name"">${esc(m[0])} <span class=""machine-total"">"
    html = html & "(${m[2][dates.length-1]})</span></span></button>`).join('')}</div></section>`).join('');"
    html = html & "    });"
    html = html & "    document.querySelectorAll('[data-machine]').forEach(b => b.onclick = () => {"
    html = html & "      selected = +b.dataset.machine; day = dates.length - 1;"
    html = html & "      lists(); chart(); logs();"
    html = html & "    });"
    html = html & "  }"
    html = html & "    function niceNumber(range, round) {"
    html = html & "      const exponent = Math.floor(Math.log10(range || 1)),"
    html = html & "        fraction = range / Math.pow(10, exponent);"
    html = html & "      let nice;"
    html = html & "      if (round) {"
    html = html & "        if (fraction < 1.5) nice = 1;"
    html = html & "        else if (fraction < 3) nice = 2;"
    html = html & "        else if (fraction < 7) nice = 5;"
    html = html & "        else nice = 10"
    html = html & "      } else {"
    html = html & "        if (fraction <= 1) nice = 1;"
    html = html & "        else if (fraction <= 2) nice = 2;"
    html = html & "        else if (fraction <= 5) nice = 5;"
    html = html & "        else nice = 10"
    html = html & "      }"
    html = html & "      return nice * Math.pow(10, exponent)"
    html = html & "    }"
    html = html & "    function niceScale(maxVal) {"
    html = html & "      maxVal = Math.max(1, Math.ceil(maxVal || 0));"
    html = html & "      const range = niceNumber(maxVal, false),"
    html = html & "        step = Math.max(1, Math.ceil(niceNumber(range / 4, true))),"
    html = html & "        max = Math.max(1, Math.ceil(maxVal / step) * step);"
    html = html & "      return {"
    html = html & "        max,"
    html = html & "        step"
    html = html & "      }"
    html = html & "    }"
    html = html & "    function chart() {"
    html = html & "      const machine = machines[selected],"
    html = html & "        w = Math.max(320, $('chart').clientWidth || 940),"
    html = html & "        h = Math.max(200, $('chart').clientHeight || 320),"
    html = html & "        m = {"
    html = html & "          t: 30,"
    html = html & "          r: w < 600 ? 82 : 115,"
    html = html & "          b: 42,"
    html = html & "          l: 48"
    html = html & "        },"
    html = html & "        plotW = w - m.l - m.r,"
    html = html & "        plotH = h - m.t - m.b,"
    html = html & "        color = '#4472c4',"
    html = html & "        scale = niceScale(Math.max(...machine[2]));"
    html = html & "      const x = i => m.l + (dates.length === 1 ? .5 : i / (dates.length - 1)) * plotW,"
    html = html & "        y = v => m.t + plotH - v / scale.max * plotH;"
    html = html & "      let s = '';"
    html = html & "      for (let v = 0; v <= scale.max; v += scale.step) {"
    html = html & "        const py = y(v);"
    html = html & "        s += `<line x1=""${m.l}"" y1=""${py}"" x2=""${m.l+plotW}"" y2=""${py}"" class=""svg-grid-line""/><text x=""${m.l-10}"" y=""${py+4}"" text-anchor=""end"" class=""svg-label"">${v}</text>`"
    html = html & "      }"
    html = html & "      s += `<line x1=""${m.l}"" y1=""${m.t}"" x2=""${m.l}"" y2=""${m.t+plotH}"" class=""svg-axis""/><line x1=""${m.l}"" y1=""${m.t+plotH}"" x2=""${m.l+plotW}"" y2=""${m.t+plotH}"" class=""svg-axis""/>`;"
    html = html & "      const pts = machine[2].map((v, i) => [x(i), y(v), v]);"
    html = html & "      s += `<path d=""${pts.map((p,i)=>(i?'L':'M')+p[0].toFixed(1)+' '+p[1].toFixed(1)).join(' ')}"" class=""svg-line"" stroke=""${color}""/>`;"
    html = html & "      pts.forEach((p, i) => {"
    html = html & "        s += `<g data-day=""${i}"" tabindex=""0"" role=""button"" aria-label=""${dates[i]}：${p[2]} 筆錯誤""><circle cx=""${p[0]}"" cy=""${p[1]}"" r=""26"" fill=""transparent"" pointer-events=""all""/><circle cx=""${p[0]}"" cy=""${p[1]}"" r=""12"" class=""svg-point-halo""/><circle cx=""${p[0]}"" cy=""${p[1]}"" r=""4"" class=""svg-point"" fill=""${color}""/><text x=""${p[0]}"" y=""${p[1]-9}"""
    html = html & " text-anchor=""middle"" class=""svg-value-label"">${p[2]}</text><text x=""${p[0]}"" y=""${m.t+plotH+24}"" text-anchor=""middle"" class=""svg-label"">${dates[i].slice(5)}</text></g>`"
    html = html & "      });"
    html = html & "      const gy = m.t + 18;"
    html = html & "      s += `<line x1=""${w-m.r+24}"" y1=""${gy-4}"" x2=""${w-m.r+46}"" y2=""${gy-4}"" stroke=""${color}"" stroke-width=""2"" stroke-linecap=""round""/><text x=""${w-m.r+52}"" y=""${gy}"" class=""svg-label"">錯誤筆數</text>`;"
    html = html & "      $('chart').setAttribute('viewBox', `0 0 ${w} ${h}`);"
    html = html & "      $('chart').innerHTML = s;"
    html = html & "      $('chart-name').textContent = machine[0];"
    html = html & "      updateChartSelection();"
    html = html & "      $('chart').querySelectorAll('[data-day]').forEach(g => {"
    html = html & "        g.onclick = () => {"
    html = html & "          day = +g.dataset.day;"
    html = html & "          logs()"
    html = html & "        };"
    html = html & "        g.onkeydown = e => {"
    html = html & "          if (e.key === 'Enter' || e.key === ' ') {"
    html = html & "            e.preventDefault();"
    html = html & "            g.onclick()"
    html = html & "          }"
    html = html & "        }"
    html = html & "      })"
    html = html & "    }"
    html = html & "    function updateChartSelection() {"
    html = html & "      $('chart').querySelectorAll('[data-day]').forEach(g => {"
    html = html & "        const active = +g.dataset.day === day;"
    html = html & "        g.classList.toggle('selected', active);"
    html = html & "        g.setAttribute('aria-pressed', String(active))"
    html = html & "      })"
    html = html & "    }"
    html = html & "  function logs() {"
    html = html & "    updateChartSelection();"
    html = html & "    const m = machines[selected];"
    html = html & "    const rows = logIndex.get(JSON.stringify([m[0].toLowerCase(), dates[day]])) || [];"
    html = html & "    $('date').textContent = dates[day];"
    html = html & "    $('prev').disabled = day === 0;"
    html = html & "    $('next').disabled = day === dates.length - 1;"
    html = html & "    $('log-count').textContent = rows.length + ' 筆';"
    html = html & "    $('log-context-details').textContent = `${m[0]} / ${dates[day]} / `;"
    html = html & "    const groups = new Map();"
    html = html & "    rows.forEach(row => {"
    html = html & "      if (!groups.has(row[2])) groups.set(row[2], new Map());"
    html = html & "      const subs = groups.get(row[2]);"
    html = html & "      if (!subs.has(row[3])) subs.set(row[3], []);"
    html = html & "      subs.get(row[3]).push(row);"
    html = html & "    });"
    html = html & "    const line = row => `<div class=""log-line""><div class=""log-text"">${esc(m[1] + ' ' + m[0] + ' | ' + row[4] + ' | ' + row[5])}</div></div>`;"
    html = html & "    $('log-scroll').innerHTML = rows.length ? [...groups].map(([category,subs]) => {"
    html = html & "      const count = [...subs.values()].reduce((n,entries) => n + entries.length,0);"
    html = html & "      return `<section class=""log-group""><h3 class=""group-title"">${esc(category)}<small>${count} 筆</small></h3>${[...subs].map(([sub,entries]) => `<h4 class=""sub-title"">${esc(sub)}<small>${entries.length} 筆</small></h4><div class=""owner-log-group""><div class=""owner-first"" ${entries.length>1?'data-toggle=""true""':''}><div class=""log-line""><div class="
    html = html & """log-text"">(${entries.length}) ${esc(m[1] + ' ' + m[0] + ' | ' + entries[0][4] + ' | ' + entries[0][5])}</div></div></div>${entries.length>1?`<div class=""owner-extra"" hidden>${entries.slice(1).map(line).join('')}</div>`:''}</div>`).join('')}</section>`;"
    html = html & "    }).join('') : '<div class=""empty"">此日沒有錯誤紀錄</div>';"
    html = html & "    $('log-scroll').querySelectorAll('[data-toggle]').forEach(first => first.ondblclick = () => {"
    html = html & "      const extra = first.nextElementSibling;"
    html = html & "      extra.hidden = !extra.hidden;"
    html = html & "      resizeFrame();"
    html = html & "    });"
    html = html & "    resizeFrame();"
    html = html & "  }"
    html = html & "  $('prev').onclick = () => {if(day>0){day--;logs();}};"
    html = html & "  $('next').onclick = () => {if(day<dates.length-1){day++;logs();}};"
    html = html & "  $('today').onclick = () => {day=dates.length-1;logs();};"
    html = html & "  function resizeFrame() {"
    html = html & "    if (window.frameElement) window.frameElement.style.height = Math.ceil(document.body.getBoundingClientRect().height + 2) + 'px';"
    html = html & "  }"
    html = html & "  if (typeof ResizeObserver !== 'undefined') {"
    html = html & "    new ResizeObserver(() => {if(selected>=0&&!$('dashboard').hidden)chart();}).observe(document.querySelector('.chart'));"
    html = html & "    new ResizeObserver(resizeFrame).observe(document.body);"
    html = html & "  }"
    html = html & "  window.addEventListener('resize', () => {if(selected>=0)chart();resizeFrame();});"
    html = html & "  resizeFrame();"
    html = html & "})();"
    OwnerScript_ = html
End Function
