Attribute VB_Name = "Module4"
'==============================================================================
' Module: modPropagation
' Purpose: Propagation engine + sim sheet generator
' Location: Insert into a standard module in the .xlsm VBA project
'==============================================================================

Option Explicit

Private Const CACHE_SHEET_NAME As String = "Cache_IMS"

Private Const C_UID As Long = 1
Private Const C_NAME As Long = 2
Private Const C_DUR As Long = 3
Private Const C_START As Long = 4
Private Const C_FINISH As Long = 5
Private Const C_PRED As Long = 6
Private Const C_SUCC As Long = 7
Private Const C_SUMMARY As Long = 8
Private Const C_MILESTONE As Long = 9
Private Const C_CONTACT As Long = 10
Private Const C_PCT As Long = 11

Private Const S_UID As Long = 1
Private Const S_CONTACT As Long = 2
Private Const S_NAME As Long = 3
Private Const S_START As Long = 4
Private Const S_FINISH As Long = 5
Private Const S_USTART As Long = 6
Private Const S_UFINISH As Long = 7

Private Type PredLink
    PredUID As Long
    LinkType As String
    LagDays As Long
End Type

Public Sub ShowSimulationForm()
    Dim frm As frmSimOptions
    
    Dim cacheWs As Worksheet
    On Error Resume Next
    Set cacheWs = ThisWorkbook.Sheets(CACHE_SHEET_NAME)
    On Error GoTo 0
    
    If cacheWs Is Nothing Then
        MsgBox "No cached IMS data found. Please run 'Cache IMS Data' first.", _
               vbExclamation, "No Cache"
        Exit Sub
    End If
    
    Set frm = New frmSimOptions
    frm.Show
    
    If frm.Cancelled Then
        Unload frm
        Exit Sub
    End If
    
    Dim sheetName As String
    Dim filterMonths As Long
    sheetName = frm.SelectedSheet
    filterMonths = frm.filterMonths
    Unload frm
    
    Call RunPropagation(sheetName, filterMonths)
End Sub

Public Sub AddSimButtonToSheet(Optional targetSheetName As String = "")
    Dim ws As Worksheet
    Dim shp As Shape
    Dim btnObj As Object
    Dim lastCol As Long
    Dim btnLeft As Double
    Dim btnTop As Double
    
    If targetSheetName = "" Then
        Set ws = ActiveSheet
    Else
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(targetSheetName)
        On Error GoTo 0
        If ws Is Nothing Then Exit Sub
    End If
    
    Select Case ws.Name
        Case "Control Panel", "Cache_IMS", "Cache_IMS_Full", "Template", "Holidays"
            Exit Sub
    End Select
    If InStr(1, ws.Name, "_Sim", vbTextCompare) > 0 Then Exit Sub
    
    ' Remove existing buttons, arrow, and label if present
    For Each shp In ws.Shapes
        If shp.Name = "btnSimStatus" Or shp.Name = "btnSimArrow" Or _
           shp.Name = "btnSimLabel" Or shp.Name = "btnSimGear" Then
            shp.Delete
        End If
    Next shp
    
    ' Find the right edge of used area
    lastCol = ws.Cells(2, ws.Columns.count).End(xlToLeft).Column
    btnLeft = ws.Cells(1, lastCol + 1).Left + 30
    btnTop = ws.Cells(1, 1).Top + 14
    
    ' "Press here" label - centered above button
    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, _
              btnLeft, btnTop - 18, 126, 14)
    With shp
        .Name = "btnSimLabel"
        .TextFrame2.TextRange.Text = "Press here"
        .TextFrame2.TextRange.Font.Size = 8
        .TextFrame2.TextRange.Font.Italic = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(180, 40, 40)
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
    End With
    
    ' Arrow pointing to button
    Dim arrowLeft As Double
    arrowLeft = ws.Cells(1, lastCol + 1).Left + 4
    
    Set shp = ws.Shapes.AddShape(msoShapeRightArrow, arrowLeft, btnTop + 4, 24, 16)
    With shp
        .Name = "btnSimArrow"
        .Fill.ForeColor.RGB = RGB(180, 40, 40)
        .Line.Visible = msoFalse
        .Shadow.Visible = msoFalse
    End With
    
    ' Main button - quick simulate (no form)
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop, 126, 30)
    With shp
        .Name = "btnSimStatus"
        .TextFrame2.TextRange.Text = "Simulate Changes"
        .TextFrame2.TextRange.Font.Size = 9.5
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.MarginLeft = 4
        .TextFrame2.MarginRight = 4
        .TextFrame2.MarginTop = 0
        .TextFrame2.MarginBottom = 0
        .Fill.ForeColor.RGB = RGB(194, 41, 41)
        .Fill.BackColor.RGB = RGB(128, 22, 22)
        .Line.ForeColor.RGB = RGB(92, 16, 16)
        .Line.Weight = 1.5
        .Shadow.Type = msoShadow14
        .Shadow.Blur = 5
        .Shadow.OffsetX = 2
        .Shadow.OffsetY = 2
        .Shadow.ForeColor.RGB = RGB(0, 0, 0)
        .Shadow.Transparency = 0.65
        .ThreeD.BevelTopType = msoBevelCircle
        .ThreeD.BevelTopDepth = 2.5
        .ThreeD.BevelTopInset = 2.5
        .OnAction = "QuickSimulate"
    End With
    
   ' Gear button - opens full options form
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft + 130, btnTop + 3, 22, 22)
    With shp
        .Name = "btnSimGear"
        .TextFrame2.TextRange.Text = ChrW(9881)
        .TextFrame2.TextRange.Font.Size = 14
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(80, 80, 80)
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.MarginLeft = 0
        .TextFrame2.MarginRight = 0
        .TextFrame2.MarginTop = 0
        .TextFrame2.MarginBottom = 0
        .Fill.ForeColor.RGB = RGB(235, 235, 235)
        .Line.ForeColor.RGB = RGB(160, 160, 160)
        .Line.Weight = 1.25
        .Shadow.Type = msoShadow14
        .Shadow.Blur = 2
        .Shadow.OffsetX = 1
        .Shadow.OffsetY = 1
        .Shadow.ForeColor.RGB = RGB(0, 0, 0)
        .Shadow.Transparency = 0.8
        .ThreeD.BevelTopType = msoBevelCircle
        .ThreeD.BevelTopDepth = 1.5
        .ThreeD.BevelTopInset = 1.5
        .OnAction = "ShowSimulationForm"
    End With
End Sub

Public Sub QuickSimulate()
    ' Quick simulate - uses active sheet, same-as-status-sheet filter, no form
    Dim cacheWs As Worksheet
    On Error Resume Next
    Set cacheWs = ThisWorkbook.Sheets("Cache_IMS")
    On Error GoTo 0
    
    If cacheWs Is Nothing Then
        MsgBox "No cached IMS data found. Please run 'Cache IMS Data' first.", _
               vbExclamation, "No Cache"
        Exit Sub
    End If
    
    Dim sheetName As String
    sheetName = ActiveSheet.Name
    
    ' Validate it's a status sheet
    Select Case sheetName
        Case "Control Panel", "Cache_IMS", "Cache_IMS_Full", "Template", "Holidays"
            MsgBox "Please navigate to a status sheet first.", vbExclamation, "Wrong Sheet"
            Exit Sub
    End Select
    
    If InStr(1, sheetName, "_Sim", vbTextCompare) > 0 Then
        MsgBox "This is a simulation sheet. Please go back to the original " & _
               "status sheet to make changes and re-simulate.", _
               vbExclamation, "Sim Sheet"
        Exit Sub
    End If
    
    ' Run propagation with default filter (same as status sheet)
    Call RunPropagation(sheetName, 0)
End Sub

Private Sub RunPropagation(statusSheetName As String, filterMonths As Long)
    Dim statusWs As Worksheet
    Dim cacheWs As Worksheet
    Dim simWs As Worksheet
    Dim t As Double
    
    t = Timer
    
    Set statusWs = ThisWorkbook.Sheets(statusSheetName)
    Set cacheWs = ThisWorkbook.Sheets(CACHE_SHEET_NAME)
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Call RefreshHolidayCache
    
    '=== STEP 1: Load cache ===
    Dim cacheData As Object
    Set cacheData = CreateObject("Scripting.Dictionary")
    
    Dim cacheArr As Variant
    Dim lastCacheRow As Long
    lastCacheRow = cacheWs.Cells(cacheWs.Rows.count, 1).End(xlUp).Row
    
    If lastCacheRow < 2 Then
        MsgBox "Cache is empty.", vbExclamation
        GoTo Cleanup
    End If
    
    cacheArr = cacheWs.Range(cacheWs.Cells(2, 1), cacheWs.Cells(lastCacheRow, 11)).Value
    
    Dim i As Long
    For i = 1 To UBound(cacheArr, 1)
        If Not IsEmpty(cacheArr(i, C_UID)) Then
            Dim uid As Long
            uid = CLng(cacheArr(i, C_UID))
            If Not cacheData.Exists(uid) Then
                Dim taskInfo(1 To 10) As Variant
                taskInfo(1) = cacheArr(i, C_NAME)
                taskInfo(2) = cacheArr(i, C_DUR)
                taskInfo(3) = cacheArr(i, C_START)
                taskInfo(4) = cacheArr(i, C_FINISH)
                taskInfo(5) = cacheArr(i, C_PRED)
                taskInfo(6) = cacheArr(i, C_SUCC)
                taskInfo(7) = cacheArr(i, C_SUMMARY)
                taskInfo(8) = cacheArr(i, C_MILESTONE)
                taskInfo(9) = cacheArr(i, C_CONTACT)
                taskInfo(10) = cacheArr(i, C_PCT)
                cacheData.Add uid, taskInfo
            End If
        End If
    Next i
    
    '=== STEP 2: Read status sheet ===
    Dim lastStatusRow As Long
    lastStatusRow = statusWs.Cells(statusWs.Rows.count, S_UID).End(xlUp).Row
    
    If lastStatusRow < 3 Then
        MsgBox "Status sheet appears empty.", vbExclamation
        GoTo Cleanup
    End If
    
    Dim statusArr As Variant
    statusArr = statusWs.Range(statusWs.Cells(3, 1), statusWs.Cells(lastStatusRow, 7)).Value
    
    Dim statusDate As Date
    On Error Resume Next
    statusDate = CDate(statusWs.Cells(1, 1).Value)
    On Error GoTo 0
    If Year(statusDate) < 2000 Then statusDate = Date
    
    Dim resolvedDates As Object
    Set resolvedDates = CreateObject("Scripting.Dictionary")
    
    Dim sheetUIDs As Object
    Set sheetUIDs = CreateObject("Scripting.Dictionary")
    
    Dim warnings As String
    warnings = ""
    
    Dim ro As Long
    For ro = 1 To UBound(statusArr, 1)
        If Not IsEmpty(statusArr(ro, S_UID)) Then
            Dim sUID As Long
            sUID = 0
            On Error Resume Next
            sUID = CLng(statusArr(ro, S_UID))
            On Error GoTo 0
            If sUID = 0 Then GoTo NextStatusRow
            
            sheetUIDs(sUID) = ro
            
            Dim taskDur As Double
            Dim origCacheStart As Variant
            Dim origCacheFinish As Variant
            taskDur = 0
            origCacheStart = Empty
            origCacheFinish = Empty
            
            If cacheData.Exists(sUID) Then
                Dim ci As Variant
                ci = cacheData(sUID)
                If ci(7) = True Then GoTo NextStatusRow
                taskDur = CDbl(ci(2))
                origCacheStart = ci(3)
                origCacheFinish = ci(4)
            End If
            
            Dim origStart As Variant, origFinish As Variant
            origStart = statusArr(ro, S_START)
            origFinish = statusArr(ro, S_FINISH)
            
            Dim updStart As Variant, updFinish As Variant
            updStart = statusArr(ro, S_USTART)
            updFinish = statusArr(ro, S_UFINISH)
            
            Dim userEnteredStart As Boolean
            Dim userEnteredFinish As Boolean

            Dim preserveStart As Boolean

            Dim preserveFinish As Boolean
            userEnteredStart = False
            userEnteredFinish = False

            preserveStart = False

            preserveFinish = False
            
            Dim effStart As Variant, effFinish As Variant
            
            If IsDate(updStart) Then
                effStart = CDate(updStart)
                If IsDate(origStart) Then
                    If Int(CDate(updStart)) <> Int(CDate(origStart)) Then
                        userEnteredStart = True
                    End If
                Else
                    userEnteredStart = True
                End If
            Else
                effStart = ResolveDateFallback(updStart, origStart, "start", sUID, warnings)
            End If
            
            If IsDate(updFinish) Then
                effFinish = CDate(updFinish)
                If IsDate(origFinish) Then
                    If Int(CDate(updFinish)) <> Int(CDate(origFinish)) Then
                        userEnteredFinish = True
                    End If
                Else
                    userEnteredFinish = True
                End If
            Else
                effFinish = ResolveDateFallback(updFinish, origFinish, "finish", sUID, warnings)
            End If
            
            If userEnteredStart And Not userEnteredFinish Then
                If IsDate(effFinish) Then

                    preserveFinish = True

                ElseIf IsDate(effStart) And taskDur > 0 Then
                    If taskDur = 1 Then
                        effFinish = effStart
                    Else
                        effFinish = AddWorkdays(CDate(effStart), CLng(taskDur) - 1)
                    End If
                    userEnteredFinish = False
                End If
            End If
            
            If userEnteredFinish And Not userEnteredStart Then

                If IsDate(effStart) Then

                    preserveStart = True

                ElseIf IsDate(effFinish) And taskDur > 0 Then

                    If taskDur = 1 Then

                        effStart = effFinish

                    Else

                        effStart = AddWorkdays(CDate(effFinish), -CLng(taskDur) + 1)

                    End If

                    userEnteredStart = False

                End If

            End If

            preserveStart = preserveStart Or userEnteredStart

            preserveFinish = preserveFinish Or userEnteredFinish

            Dim resolved(1 To 9) As Variant
            resolved(1) = effStart
            resolved(2) = effFinish
            resolved(3) = preserveStart
            resolved(4) = preserveFinish
            resolved(5) = taskDur
            resolved(6) = origStart
            resolved(7) = origFinish

            resolved(8) = userEnteredStart

            resolved(9) = userEnteredFinish
            resolvedDates(sUID) = resolved
        End If
NextStatusRow:
    Next ro
    
    '=== STEP 3: Expand scope ===
    Dim contactName As String
    contactName = ""
    For ro = 1 To UBound(statusArr, 1)
        If Len(Trim(CStr(statusArr(ro, S_CONTACT) & ""))) > 0 Then
            contactName = Trim(CStr(statusArr(ro, S_CONTACT)))
            Exit For
        End If
    Next ro
    
    Dim extraUIDs As Object
    Set extraUIDs = CreateObject("Scripting.Dictionary")
    
    If filterMonths <> 0 And contactName <> "" Then
        Dim cutoffDate As Date
        If filterMonths = -1 Then
            cutoffDate = DateSerial(2099, 12, 31)
        Else
            cutoffDate = DateAdd("m", filterMonths, statusDate)
        End If
        
        Dim k As Variant
        For Each k In cacheData.Keys
            Dim cTask As Variant
            cTask = cacheData(k)
            If LCase(Trim(CStr(cTask(9) & ""))) = LCase(Trim(contactName)) Then
                If Not sheetUIDs.Exists(CLng(k)) Then
                    If IsDate(cTask(4)) Then
                        If CDate(cTask(4)) <= cutoffDate Then
                            extraUIDs(CLng(k)) = True
                            Dim extraRes(1 To 7) As Variant
                            extraRes(1) = cTask(3)
                            extraRes(2) = cTask(4)
                            extraRes(3) = False
                            extraRes(4) = False
                            extraRes(5) = CDbl(cTask(2))
                            extraRes(6) = cTask(3)
                            extraRes(7) = cTask(4)
                            resolvedDates(CLng(k)) = extraRes
                        End If
                    End If
                End If
            End If
        Next k
    End If
    
    '=== STEP 4: Propagation ===
    Dim allUIDs As Object
    Set allUIDs = CreateObject("Scripting.Dictionary")
    For Each k In sheetUIDs.Keys
        allUIDs(k) = True
    Next k
    For Each k In extraUIDs.Keys
        allUIDs(k) = True
    Next k
    
    Dim changed As Boolean
    Dim iterations As Long
    iterations = 0
    
    Do
        changed = False
        iterations = iterations + 1
        
        For Each k In allUIDs.Keys
            Dim thisUID As Long
            thisUID = CLng(k)
            
            If Not cacheData.Exists(thisUID) Then GoTo NextPropTask
            
            Dim thisTask As Variant
            thisTask = cacheData(thisUID)
            If thisTask(7) = True Then GoTo NextPropTask
            
            Dim predStr As String
            predStr = CStr(thisTask(5) & "")
            If Len(predStr) = 0 Then GoTo NextPropTask
            
            Dim predLinks() As PredLink
            Dim predCount As Long
            Call ParsePredecessors(predStr, predLinks, predCount)
            If predCount = 0 Then GoTo NextPropTask
            
            If Not resolvedDates.Exists(thisUID) Then GoTo NextPropTask
            
            Dim curRes As Variant
            curRes = resolvedDates(thisUID)
            
            Dim curStart As Variant, curFinish As Variant
            Dim dur As Double
            curStart = curRes(1)
            curFinish = curRes(2)
            dur = CDbl(curRes(5))
            
            Dim drivenStart As Date
            Dim drivenFinish As Date
            Dim hasDrivenStart As Boolean
            Dim hasDrivenFinish As Boolean
            hasDrivenStart = False
            hasDrivenFinish = False
            
            Dim p As Long
            For p = 1 To predCount
                Dim pUID As Long
                pUID = predLinks(p).PredUID
                
                Dim predStart As Variant, predFinish As Variant
                
                If resolvedDates.Exists(pUID) Then
                    Dim predRes As Variant
                    predRes = resolvedDates(pUID)
                    predStart = predRes(1)
                    predFinish = predRes(2)
                ElseIf cacheData.Exists(pUID) Then
                    Dim predCache As Variant
                    predCache = cacheData(pUID)
                    predStart = predCache(3)
                    predFinish = predCache(4)
                Else
                    GoTo NextPred
                End If
                
                Dim calcDate As Date
                
                Select Case predLinks(p).LinkType
                    Case "FS"
                        If IsDate(predFinish) Then
                            If predLinks(p).LagDays = 0 Then
                                calcDate = NextWorkday(CDate(predFinish) + 1)
                            Else
                                calcDate = AddWorkdays(CDate(predFinish), predLinks(p).LagDays)
                                calcDate = NextWorkday(calcDate)
                            End If
                            If Not hasDrivenStart Then
                                drivenStart = calcDate
                                hasDrivenStart = True
                            ElseIf calcDate > drivenStart Then
                                drivenStart = calcDate
                            End If
                        End If
                    Case "SS"
                        If IsDate(predStart) Then
                            If predLinks(p).LagDays = 0 Then
                                calcDate = NextWorkday(CDate(predStart))
                            Else
                                calcDate = AddWorkdays(CDate(predStart), predLinks(p).LagDays)
                                calcDate = NextWorkday(calcDate)
                            End If
                            If Not hasDrivenStart Then
                                drivenStart = calcDate
                                hasDrivenStart = True
                            ElseIf calcDate > drivenStart Then
                                drivenStart = calcDate
                            End If
                        End If
                    Case "FF"
                        If IsDate(predFinish) Then
                            If predLinks(p).LagDays = 0 Then
                                calcDate = PrevWorkday(CDate(predFinish))
                            Else
                                calcDate = AddWorkdays(CDate(predFinish), predLinks(p).LagDays)
                                calcDate = PrevWorkday(calcDate)
                            End If
                            If Not hasDrivenFinish Then
                                drivenFinish = calcDate
                                hasDrivenFinish = True
                            ElseIf calcDate > drivenFinish Then
                                drivenFinish = calcDate
                            End If
                        End If
                    Case "SF"
                        If IsDate(predStart) Then
                            If predLinks(p).LagDays = 0 Then
                                calcDate = PrevWorkday(CDate(predStart))
                            Else
                                calcDate = AddWorkdays(CDate(predStart), predLinks(p).LagDays)
                                calcDate = PrevWorkday(calcDate)
                            End If
                            If Not hasDrivenFinish Then
                                drivenFinish = calcDate
                                hasDrivenFinish = True
                            ElseIf calcDate > drivenFinish Then
                                drivenFinish = calcDate
                            End If
                        End If
                End Select
NextPred:
            Next p
            
            Dim newStart As Variant
            Dim newFinish As Variant
            Dim taskChanged As Boolean
            newStart = curStart
            newFinish = curFinish
            taskChanged = False
            
            If hasDrivenStart And curRes(3) = False Then
                If IsDate(curStart) Then
                    If drivenStart > CDate(curStart) Then
                        newStart = drivenStart
                        taskChanged = True
                    End If
                Else
                    newStart = drivenStart
                    taskChanged = True
                End If
            End If
            
            If hasDrivenFinish And curRes(4) = False Then
                If IsDate(newFinish) Then
                    If drivenFinish > CDate(newFinish) Then
                        newFinish = drivenFinish
                        taskChanged = True
                    End If
                Else
                    newFinish = drivenFinish
                    taskChanged = True
                End If
            End If
            
            If taskChanged And IsDate(newStart) And curRes(4) = False Then
                If dur > 0 Then
                    If dur = 1 Then
                        newFinish = newStart
                    Else
                        newFinish = AddWorkdays(CDate(newStart), CLng(dur) - 1)
                    End If
                End If
                If hasDrivenFinish Then
                    If IsDate(newFinish) And drivenFinish > CDate(newFinish) Then
                        newFinish = drivenFinish
                    End If
                End If
            End If
            
            If taskChanged And IsDate(newFinish) And curRes(3) = False Then

                If (Not IsDate(newStart)) Or (hasDrivenStart = False) Then

                    If dur > 0 Then

                        If dur = 1 Then

                            newStart = newFinish

                        Else

                            newStart = AddWorkdays(CDate(newFinish), -CLng(dur) + 1)

                        End If

                    End If

                End If

            End If

            If taskChanged Then
                Dim updRes(1 To 9) As Variant
                updRes(1) = newStart
                updRes(2) = newFinish
                updRes(3) = curRes(3)
                updRes(4) = curRes(4)
                updRes(5) = dur
                updRes(6) = curRes(6)
                updRes(7) = curRes(7)

                updRes(8) = curRes(8)

                updRes(9) = curRes(9)
                resolvedDates(thisUID) = updRes
                changed = True
            End If
NextPropTask:
        Next k
    Loop While changed And iterations < 200
    
    '=== STEP 5: Generate sim sheet ===
    Dim simName As String
    simName = GetNextSimName(statusSheetName)
    
    statusWs.Copy After:=statusWs
    Set simWs = ActiveSheet
    simWs.Name = simName
    
    ' Remove buttons from sim sheet (read-only)
    Dim delShp As Shape
    For Each delShp In simWs.Shapes
        If delShp.Name = "btnSimStatus" Or delShp.Name = "btnCacheIMS" Or _
           delShp.Name = "btnReset" Or delShp.Name = "btnSimulate" Or _
           delShp.Name = "btnSimArrow" Or delShp.Name = "btnSimLabel" Or _
           delShp.Name = "btnSimGear" Then
            delShp.Delete
        End If
    Next delShp
    
    simWs.Cells(1, 1).Value = "SIMULATION"
    simWs.Cells(1, 1).Font.Color = RGB(192, 0, 0)
    simWs.Cells(1, 1).Font.Bold = True
    
    Dim noteText As String
    noteText = "Generated: " & Format(Now, "mm/dd/yyyy hh:nn:ss") & " | Filter: "
    Select Case filterMonths
        Case 0: noteText = noteText & "Same as Status Sheet"
        Case -1: noteText = noteText & "All Tasks"
        Case Else: noteText = noteText & "Next " & filterMonths & " Months"
    End Select
    On Error Resume Next
    simWs.Cells(1, 1).Comment.Delete
    On Error GoTo 0
    simWs.Cells(1, 1).AddComment noteText
    
    '=== STEP 6: Write dates with proper formatting ===
    Dim lastSimRow As Long
    lastSimRow = simWs.Cells(simWs.Rows.count, S_UID).End(xlUp).Row
    
    Dim darkRedFill As Long: darkRedFill = RGB(220, 100, 100)
    Dim lightRedFill As Long: lightRedFill = RGB(255, 180, 180)
    Dim darkGreenFill As Long: darkGreenFill = RGB(100, 180, 100)
    Dim lightGreenFill As Long: lightGreenFill = RGB(198, 239, 206)
    Dim yellowFill As Long: yellowFill = RGB(255, 255, 0)
    Dim grayFont As Long: grayFont = RGB(180, 180, 180)
    Dim blackFont As Long: blackFont = RGB(0, 0, 0)
    Dim redFont As Long: redFont = RGB(255, 0, 0)
    
    For ro = 3 To lastSimRow
        Dim simUID As Long
        simUID = 0
        On Error Resume Next
        simUID = CLng(simWs.Cells(ro, S_UID).Value)
        On Error GoTo 0
        If simUID = 0 Then GoTo NextSimRow
        
        If resolvedDates.Exists(simUID) Then
            Dim simRes As Variant
            simRes = resolvedDates(simUID)
            
            Dim simOrigStart As Variant: simOrigStart = simRes(6)
            Dim simOrigFinish As Variant: simOrigFinish = simRes(7)
            Dim userStart As Boolean: userStart = CBool(simRes(8))
            Dim userFinish As Boolean: userFinish = CBool(simRes(9))
            
            Dim uStartRaw As String: uStartRaw = LCase(Trim(CStr(simWs.Cells(ro, S_USTART).Value & "")))
            Dim uFinishRaw As String: uFinishRaw = LCase(Trim(CStr(simWs.Cells(ro, S_UFINISH).Value & "")))
            
            Dim isStartShould As Boolean
            isStartShould = (uStartRaw = "should strt" Or uStartRaw = "should start" Or uStartRaw = "should str")
            Dim isFinishShould As Boolean
            isFinishShould = (uFinishRaw = "should fin" Or uFinishRaw = "should finish")
            
            If isStartShould And Not IsDate(simRes(1)) Then
                simWs.Cells(ro, S_USTART).Interior.Color = yellowFill
                simWs.Cells(ro, S_USTART).Font.Color = redFont
            Else
                If IsDate(simRes(1)) Then
                    simWs.Cells(ro, S_USTART).Value = CDate(simRes(1))
                    simWs.Cells(ro, S_USTART).NumberFormat = "mm/dd/yyyy"
                End If
                
                Dim startDelta As Long: startDelta = 0
                If IsDate(simRes(1)) And IsDate(simOrigStart) Then
                    If Int(CDate(simRes(1))) > Int(CDate(simOrigStart)) Then
                        startDelta = 1
                    ElseIf Int(CDate(simRes(1))) < Int(CDate(simOrigStart)) Then
                        startDelta = -1
                    End If
                End If
                
                If startDelta = 1 Then
                    If userStart Then
                        simWs.Cells(ro, S_USTART).Interior.Color = darkRedFill
                    Else
                        simWs.Cells(ro, S_USTART).Interior.Color = lightRedFill
                    End If
                    simWs.Cells(ro, S_USTART).Font.Color = blackFont
                ElseIf startDelta = -1 Then
                    If userStart Then
                        simWs.Cells(ro, S_USTART).Interior.Color = darkGreenFill
                    Else
                        simWs.Cells(ro, S_USTART).Interior.Color = lightGreenFill
                    End If
                    simWs.Cells(ro, S_USTART).Font.Color = blackFont
                Else
                    simWs.Cells(ro, S_USTART).Interior.ColorIndex = xlNone
                    If IsDate(simRes(1)) Then

                        simWs.Cells(ro, S_USTART).Font.Color = blackFont

                    Else

                        simWs.Cells(ro, S_USTART).Font.Color = grayFont

                    End If
                End If
            End If
            
            If isFinishShould And Not IsDate(simRes(2)) Then
                simWs.Cells(ro, S_UFINISH).Interior.Color = yellowFill
                simWs.Cells(ro, S_UFINISH).Font.Color = redFont
            Else
                If IsDate(simRes(2)) Then
                    simWs.Cells(ro, S_UFINISH).Value = CDate(simRes(2))
                    simWs.Cells(ro, S_UFINISH).NumberFormat = "mm/dd/yyyy"
                End If
                
                Dim finishDelta As Long: finishDelta = 0
                If IsDate(simRes(2)) And IsDate(simOrigFinish) Then
                    If Int(CDate(simRes(2))) > Int(CDate(simOrigFinish)) Then
                        finishDelta = 1
                    ElseIf Int(CDate(simRes(2))) < Int(CDate(simOrigFinish)) Then
                        finishDelta = -1
                    End If
                End If
                
                If finishDelta = 1 Then
                    If userFinish Then
                        simWs.Cells(ro, S_UFINISH).Interior.Color = darkRedFill
                    Else
                        simWs.Cells(ro, S_UFINISH).Interior.Color = lightRedFill
                    End If
                    simWs.Cells(ro, S_UFINISH).Font.Color = blackFont
                ElseIf finishDelta = -1 Then
                    If userFinish Then
                        simWs.Cells(ro, S_UFINISH).Interior.Color = darkGreenFill
                    Else
                        simWs.Cells(ro, S_UFINISH).Interior.Color = lightGreenFill
                    End If
                    simWs.Cells(ro, S_UFINISH).Font.Color = blackFont
                Else
                    simWs.Cells(ro, S_UFINISH).Interior.ColorIndex = xlNone
                    If IsDate(simRes(2)) Then

                        simWs.Cells(ro, S_UFINISH).Font.Color = blackFont

                    Else

                        simWs.Cells(ro, S_UFINISH).Font.Color = grayFont

                    End If
                End If
            End If
        End If
NextSimRow:
    Next ro
    
    '=== STEP 7: Clear SSI formatting and legend ===
    Dim clearRow As Long
    For clearRow = 3 To lastSimRow
        simWs.Cells(clearRow, S_START).Interior.ColorIndex = xlNone
        simWs.Cells(clearRow, S_FINISH).Interior.ColorIndex = xlNone
        simWs.Cells(clearRow, S_START).Font.Color = RGB(0, 0, 0)
        simWs.Cells(clearRow, S_FINISH).Font.Color = RGB(0, 0, 0)
    Next clearRow
    
    Dim scanRow As Long
    Dim legendStartRow As Long
    legendStartRow = 0
    For scanRow = lastSimRow + 1 To lastSimRow + 20
        Dim cellText As String
        cellText = LCase(Trim(CStr(simWs.Cells(scanRow, 3).Value & "")))
        If InStr(1, cellText, "legend", vbTextCompare) > 0 Then
            legendStartRow = scanRow
            Exit For
        End If
        cellText = LCase(Trim(CStr(simWs.Cells(scanRow, 2).Value & "")))
        If InStr(1, cellText, "legend", vbTextCompare) > 0 Or _
           InStr(1, cellText, "gray", vbTextCompare) > 0 Then
            legendStartRow = scanRow
            Exit For
        End If
    Next scanRow
    
    If legendStartRow > 0 Then
        simWs.Range(simWs.Cells(legendStartRow, 1), _
                     simWs.Cells(legendStartRow + 10, 7)).Clear
    End If
    
    '=== STEP 7.5: Match column widths and row heights ===
    Dim c As Long
    For c = 1 To 7
        simWs.Columns(c).ColumnWidth = statusWs.Columns(c).ColumnWidth
    Next c
    
    Dim rw As Long
    For rw = 1 To lastSimRow + 10
        simWs.Rows(rw).RowHeight = statusWs.Rows(rw).RowHeight
    Next rw
    
    '=== STEP 7.6: Create legend ===

    Call BuildSimulationLegend(simWs, simWs.Cells(3, 8).Left + 10, simWs.Cells(3, 8).Top)
    GoTo LegendFinished
    Dim legendLeft As Double
    Dim legendTop As Double
    legendLeft = simWs.Cells(3, 8).Left + 10
    legendTop = simWs.Cells(3, 8).Top
    
    Dim shp As Shape
    Set shp = simWs.Shapes.AddShape(msoShapeRoundedRectangle, _
              legendLeft, legendTop, 230, 235)
    
    With shp
        .Name = "simLegend"
        .Fill.ForeColor.RGB = RGB(252, 252, 252)
        .Fill.Transparency = 0
        .Line.ForeColor.RGB = RGB(200, 200, 200)
        .Line.Weight = 0.75
        .Shadow.Type = msoShadow21
        .Shadow.Blur = 8
        .Shadow.Transparency = 0.75
        
        With .TextFrame2
            .MarginLeft = 28
            .MarginRight = 10
            .MarginTop = 8
            .MarginBottom = 8
            .WordWrap = msoTrue
        End With
    End With
    
    Dim tf As Object
    Set tf = shp.TextFrame2.TextRange
    tf.Text = ""
    
    Dim lines(1 To 15) As String
    Dim lineColors(1 To 15) As Long
    Dim lineSizes(1 To 15) As Long
    Dim lineBold(1 To 15) As Boolean
    Dim lineItalic(1 To 15) As Boolean
    
    lines(1) = "Legend"
    lineColors(1) = RGB(40, 40, 40): lineSizes(1) = 11: lineBold(1) = True: lineItalic(1) = False
    lines(2) = "- No Change or Effect"
    lineColors(2) = RGB(170, 170, 170): lineSizes(2) = 8: lineBold(2) = False: lineItalic(2) = False
    lines(3) = " "
    lineColors(3) = RGB(0, 0, 0): lineSizes(3) = 3: lineBold(3) = False: lineItalic(3) = False
    lines(4) = "Bad Change"
    lineColors(4) = RGB(40, 40, 40): lineSizes(4) = 9: lineBold(4) = True: lineItalic(4) = False
    lines(5) = "- Task's Changed (Dates are pushed out)"
    lineColors(5) = RGB(220, 100, 100): lineSizes(5) = 8: lineBold(5) = False: lineItalic(5) = False
    lines(6) = "- Its Cascading Effect"
    lineColors(6) = RGB(240, 150, 150): lineSizes(6) = 8: lineBold(6) = False: lineItalic(6) = False
    lines(7) = " "
    lineColors(7) = RGB(0, 0, 0): lineSizes(7) = 3: lineBold(7) = False: lineItalic(7) = False
    lines(8) = "Good Change"
    lineColors(8) = RGB(40, 40, 40): lineSizes(8) = 9: lineBold(8) = True: lineItalic(8) = False
    lines(9) = "- Task's Changed (Dates are earlier than before)"
    lineColors(9) = RGB(80, 160, 80): lineSizes(9) = 8: lineBold(9) = False: lineItalic(9) = False
    lines(10) = "- Its Cascading Effect"
    lineColors(10) = RGB(130, 200, 130): lineSizes(10) = 8: lineBold(10) = False: lineItalic(10) = False
    lines(11) = " "
    lineColors(11) = RGB(0, 0, 0): lineSizes(11) = 3: lineBold(11) = False: lineItalic(11) = False
    lines(12) = "- Update Still Required"
    lineColors(12) = RGB(190, 140, 0): lineSizes(12) = 8: lineBold(12) = False: lineItalic(12) = False
    lines(13) = " "
    lineColors(13) = RGB(0, 0, 0): lineSizes(13) = 3: lineBold(13) = False: lineItalic(13) = False
    lines(14) = "This is a READ-ONLY simulation."
    lineColors(14) = RGB(140, 140, 140): lineSizes(14) = 7: lineBold(14) = False: lineItalic(14) = True
    lines(15) = "Update the original status sheet."
    lineColors(15) = RGB(140, 140, 140): lineSizes(15) = 7: lineBold(15) = False: lineItalic(15) = True
    
    Dim fullText As String
    Dim lineIdx As Long
    fullText = ""
    For lineIdx = 1 To 15
        If lineIdx > 1 Then fullText = fullText & vbLf
        fullText = fullText & lines(lineIdx)
    Next lineIdx
    
    tf.Text = fullText
    
    Dim charPos As Long
    charPos = 1
    
    For lineIdx = 1 To 15
        Dim lineLen As Long
        lineLen = Len(lines(lineIdx))
        If lineLen > 0 Then
            Dim charRange As Object
            Set charRange = tf.Characters(charPos, lineLen)
            charRange.Font.Size = lineSizes(lineIdx)
            charRange.Font.Fill.ForeColor.RGB = lineColors(lineIdx)
            If lineBold(lineIdx) Then
                charRange.Font.Bold = msoTrue
            Else
                charRange.Font.Bold = msoFalse
            End If
            If lineItalic(lineIdx) Then
                charRange.Font.Italic = msoTrue
            Else
                charRange.Font.Italic = msoFalse
            End If
        End If
        charPos = charPos + lineLen + 1
    Next lineIdx
    
    Dim swatchLeft As Double
    Dim swatchSize As Double
    swatchLeft = legendLeft + 10
    swatchSize = 10
    
    Dim swatchTop As Double
    Dim swt As Shape
    
    swatchTop = legendTop + 36
    Set swt = simWs.Shapes.AddShape(msoShapeRoundedRectangle, swatchLeft, swatchTop, swatchSize, swatchSize)
    swt.Fill.ForeColor.RGB = RGB(200, 200, 200)
    swt.Line.Visible = msoFalse
    swt.Name = "swatchGray"
    
    swatchTop = legendTop + 76
    Set swt = simWs.Shapes.AddShape(msoShapeRoundedRectangle, swatchLeft, swatchTop, swatchSize, swatchSize)
    swt.Fill.ForeColor.RGB = RGB(220, 100, 100)
    swt.Line.Visible = msoFalse
    swt.Name = "swatchDarkRed"
    
    swatchTop = legendTop + 90
    Set swt = simWs.Shapes.AddShape(msoShapeRoundedRectangle, swatchLeft, swatchTop, swatchSize, swatchSize)
    swt.Fill.ForeColor.RGB = RGB(255, 180, 180)
    swt.Line.Visible = msoFalse
    swt.Name = "swatchLightRed"
    
    swatchTop = legendTop + 128
    Set swt = simWs.Shapes.AddShape(msoShapeRoundedRectangle, swatchLeft, swatchTop, swatchSize, swatchSize)
    swt.Fill.ForeColor.RGB = RGB(80, 160, 80)
    swt.Line.Visible = msoFalse
    swt.Name = "swatchDarkGreen"
    
    swatchTop = legendTop + 142
    Set swt = simWs.Shapes.AddShape(msoShapeRoundedRectangle, swatchLeft, swatchTop, swatchSize, swatchSize)
    swt.Fill.ForeColor.RGB = RGB(180, 230, 180)
    swt.Line.Visible = msoFalse
    swt.Name = "swatchLightGreen"
    
    swatchTop = legendTop + 170
    Set swt = simWs.Shapes.AddShape(msoShapeRoundedRectangle, swatchLeft, swatchTop, swatchSize, swatchSize)
    swt.Fill.ForeColor.RGB = RGB(240, 210, 50)
    swt.Line.Visible = msoFalse
    swt.Name = "swatchYellow"
    
LegendFinished:

    '=== STEP 8: Protect and finish ===
    simWs.Protect Password:="", UserInterfaceOnly:=True
    simWs.Activate
    
    Dim elapsed As Double
    elapsed = Timer - t
    
    Dim msg As String
    msg = "Simulation complete!" & vbCrLf & _
          "Sheet: " & simName & vbCrLf & _
          "Iterations: " & iterations & vbCrLf & _
          "Time: " & Format(elapsed, "0.0") & " seconds"
    
    If Len(warnings) > 0 Then
        msg = msg & vbCrLf & vbCrLf & _
              "WARNINGS:" & vbCrLf & warnings
    End If
    
    MsgBox msg, IIf(Len(warnings) > 0, vbExclamation, vbInformation), "Simulation"
    
Cleanup:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Private Function ResolveDateFallback(updVal As Variant, origVal As Variant, _
                                     fieldName As String, uid As Long, _
                                     ByRef warnings As String) As Variant
    Dim strVal As String
    strVal = LCase(Trim(CStr(updVal & "")))
    
    If strVal = "" Or strVal = "should strt" Or strVal = "should fin" Or _
       strVal = "should start" Or strVal = "should finish" Or _
       strVal = "should str" Then
        If IsDate(origVal) Then
            ResolveDateFallback = CDate(origVal)
        Else
            ResolveDateFallback = Empty
        End If
        Exit Function
    End If
    
    warnings = warnings & "UID " & uid & ": Non-date text '" & _
               CStr(updVal) & "' in Updated " & fieldName & " column" & vbCrLf
    
    If IsDate(origVal) Then
        ResolveDateFallback = CDate(origVal)
    Else
        ResolveDateFallback = Empty
    End If
End Function

Private Sub ParsePredecessors(predStr As String, ByRef links() As PredLink, ByRef count As Long)
    Dim parts() As String
    count = 0
    If Len(Trim(predStr)) = 0 Then Exit Sub
    
    parts = Split(predStr, ",")
    ReDim links(1 To UBound(parts) + 1)
    
    Dim idx As Long
    For idx = 0 To UBound(parts)
        Dim token As String
        token = Trim(CStr(parts(idx)))
        If Len(token) = 0 Then GoTo NextToken
        
        token = Replace(token, " days", "", , , vbTextCompare)
        token = Replace(token, " day", "", , , vbTextCompare)
        If Right(token, 1) = "d" Or Right(token, 1) = "D" Then
            If Len(token) > 1 Then
                If IsNumeric(Mid(token, Len(token) - 1, 1)) Or _
                   Mid(token, Len(token) - 1, 1) = "+" Or _
                   Mid(token, Len(token) - 1, 1) = "-" Then
                    token = Left(token, Len(token) - 1)
                End If
            End If
        End If
        If Right(token, 1) = "e" Or Right(token, 1) = "E" Then
            token = Left(token, Len(token) - 1)
        End If
        
        count = count + 1
        
        Dim pUID As Long, pType As String, pLag As Long
        pType = "FS": pLag = 0
        
        Dim numEnd As Long: numEnd = 0
        Dim ch As String
        Dim ci2 As Long
        For ci2 = 1 To Len(token)
            ch = Mid(token, ci2, 1)
            If IsNumeric(ch) Then
                numEnd = ci2
            Else
                Exit For
            End If
        Next ci2
        
        If numEnd = 0 Then
            count = count - 1
            GoTo NextToken
        End If
        
        pUID = CLng(Left(token, numEnd))
        
        If numEnd < Len(token) Then
            Dim remainder As String
            remainder = Mid(token, numEnd + 1)
            If Len(remainder) >= 2 Then
                Dim lt As String
                lt = UCase(Left(remainder, 2))
                If lt = "FS" Or lt = "FF" Or lt = "SS" Or lt = "SF" Then
                    pType = lt
                    remainder = Mid(remainder, 3)
                End If
            End If
            If Len(remainder) > 0 Then
                Dim lagStr As String: lagStr = ""
                Dim li As Long
                For li = 1 To Len(remainder)
                    ch = Mid(remainder, li, 1)
                    If ch = "+" Or ch = "-" Or IsNumeric(ch) Then lagStr = lagStr & ch
                Next li
                If Len(lagStr) > 0 And IsNumeric(lagStr) Then pLag = CLng(lagStr)
            End If
        End If
        
        links(count).PredUID = pUID
        links(count).LinkType = pType
        links(count).LagDays = pLag
NextToken:
    Next idx
End Sub

Private Function GetNextSimName(baseSheet As String) As String
    Dim simNum As Long
    simNum = 1
    Do
        Dim testName As String
        testName = baseSheet & "_Sim" & simNum
        Dim ws As Worksheet
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(testName)
        On Error GoTo 0
        If ws Is Nothing Then
            GetNextSimName = testName
            Exit Function
        End If
        simNum = simNum + 1
    Loop
End Function

Public Sub ResetStatusSheet()
    Dim ws As Worksheet
    Dim simSheets As String
    simSheets = ""
    
    For Each ws In ThisWorkbook.Sheets
        If InStr(1, ws.Name, "_Sim", vbTextCompare) > 0 Then
            simSheets = simSheets & ws.Name & vbCrLf
        End If
    Next ws
    
    If simSheets = "" Then
        MsgBox "No simulation sheets found.", vbInformation, "Reset"
        Exit Sub
    End If
    
    If MsgBox("Delete ALL simulation sheets?" & vbCrLf & vbCrLf & _
              simSheets & vbCrLf & _
              "This cannot be undone.", _
              vbYesNo + vbExclamation, "Confirm Delete") = vbYes Then
        Application.DisplayAlerts = False
        For Each ws In ThisWorkbook.Sheets
            If InStr(1, ws.Name, "_Sim", vbTextCompare) > 0 Then
                ws.Delete
            End If
        Next ws
        Application.DisplayAlerts = True
        MsgBox "All simulation sheets deleted.", vbInformation, "Reset Complete"
    End If
End Sub

