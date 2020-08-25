VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub NextGen_Click()
    GameLoop.GameStep
End Sub

Private Sub Reset_Click()
    
    'Clear current option button selection
    SimkinGliderGun.Value = False
    Pentadecathlon.Value = False
    Pulsar.Value = False
    SpaceshipFleet.Value = False
    Random.Value = False
    
    'reset the board
    GameLoop.Reset
End Sub

Private Sub SimkinGliderGun_Click()

    GameLoop.CopyPreset ("SimkinGliderGun")
  
End Sub

Private Sub Pentadecathlon_Click()

    GameLoop.CopyPreset ("Pentadecathlon")
   
End Sub

Private Sub Pulsar_Click()

    GameLoop.CopyPreset ("Pulsar")
    
End Sub

Private Sub SpaceshipFleet_Click()

    GameLoop.CopyPreset ("SpaceshipFleet")
    
End Sub

Private Sub Start_Click()
    ' Use a public variable to control if clicking the
    ' button starts or stops the macro
    StopMacros = Not StopMacros
    ' change the text on the button
    Start.Caption = IIf(StopMacros, "Start", "Stop")
    ' set a condional on the color of the button
    If Start.Caption = "Stop" Then
        Start.BackColor = &HFF&
    Else
        Start.BackColor = &HC000&
    End If
    If Not StopMacros Then Game_Loop
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    '''
    ' This sets the workbook to change the values in the field by only cliking on them
    '''
    ' Make sure only one cell is selected so the code doesn't break
    If Selection.Count = 1 Then
        ' Make sure it only happens in the defined range
        If Not Intersect(Target, Range("C3:AP42")) Is Nothing Then
            ' Change it from alive to dead and vice versa each time a cell is selected
            If Selection.Value = 0 Then
                Selection.Value = 1
            Else
                Selection.Value = 0
            End If
        End If
    End If
End Sub

