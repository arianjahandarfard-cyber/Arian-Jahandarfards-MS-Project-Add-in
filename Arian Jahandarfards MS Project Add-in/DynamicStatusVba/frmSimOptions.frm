Attribute VB_Name = "frmSimOptions"
Attribute VB_Base = "0{25099E15-1D58-4B04-9197-D6A13AAD300B}{046E3BC6-31E3-419A-9683-13A99BD4CD7A}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
'==============================================================================
' UserForm: frmSimOptions
' Purpose: Lets user pick filter scope and target status sheet for simulation
'==============================================================================

Option Explicit

Public SelectedSheet As String
Public filterMonths As Long   ' 0=same as sheet, 1/2/3/6=months, -1=all tasks
Public Cancelled As Boolean

Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    
    Cancelled = True
    SelectedSheet = ""
    filterMonths = 0
    
    ' Populate combo with status sheets (exclude system sheets)
    For Each ws In ThisWorkbook.Sheets
        Select Case ws.Name
            Case "Control Panel", "Cache_IMS", "Cache_IMS_Full", "Template"
                ' Skip system sheets
            Case Else
                ' Skip sim sheets
                If InStr(1, ws.Name, "_Sim", vbTextCompare) = 0 Then
                    cboSheet.AddItem ws.Name
                End If
        End Select
    Next ws
    
    ' Default to active sheet if it's a status sheet
    Dim i As Long
    For i = 0 To cboSheet.ListCount - 1
        If cboSheet.List(i) = ActiveSheet.Name Then
            cboSheet.ListIndex = i
            Exit For
        End If
    Next i
    
    ' If nothing matched, select first item
    If cboSheet.ListIndex = -1 And cboSheet.ListCount > 0 Then
        cboSheet.ListIndex = 0
    End If
    
    optSameAsSheet.Value = True
End Sub

Private Sub btnRun_Click()
    If cboSheet.ListIndex = -1 Then
        MsgBox "Please select a status sheet.", vbExclamation, "No Sheet Selected"
        Exit Sub
    End If
    
    SelectedSheet = cboSheet.Value
    
    If optSameAsSheet.Value Then
        filterMonths = 0
    ElseIf optNext1.Value Then
        filterMonths = 1
    ElseIf optNext2.Value Then
        filterMonths = 2
    ElseIf optNext3.Value Then
        filterMonths = 3
    ElseIf optNext6.Value Then
        filterMonths = 6
    ElseIf optAllTasks.Value Then
        filterMonths = -1
    End If
    
    Cancelled = False
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    Cancelled = True
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancelled = True
        Me.Hide
        Cancel = True
    End If
End Sub
