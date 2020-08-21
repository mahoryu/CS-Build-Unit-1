Attribute VB_Name = "GameLoop"
Option Explicit

#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
    'For 64 Bit Systems
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If
    'For 32 Bit Systems

Public StopMacros As Boolean

Sub Game_Loop()
        
    ' Keep the game running until we hit a specified generation
    While Not StopMacros
        
        DoEvents
        If StopMacros Then Exit Sub
        
        'Copy the successor generation to the clipboard and then
        '   paste it over the current generation
        Sheets("Successor Generation").Range("C3:AP42").Copy
        Sheets("Current Generation").Range("C3").PasteSpecial Paste:=xlPasteValues, _
            Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
        ' Update the generation value
        Range("AY2").Value = Range("AY2").Value + 1
        
        'Set the sleep speed to the value of the scroll bar on the sheet
        Dim speed As Integer
        ' setting it to max - value so that the faster speed is to the right side of the scroll bar
        With Sheets("Current Generation").SpeedScaler
            speed = .Max - .Value
        End With
        Sleep (speed)
        
    Wend
           
End Sub

Sub GameStep()

    'Copy the successor generation to the clipboard and then
    '   paste it over the current generation
    Sheets("Successor Generation").Range("C3:AP42").Copy
    Sheets("Current Generation").Range("C3").PasteSpecial Paste:=xlPasteValues, _
        Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    ' Update the generation value
    Range("AY2").Value = Range("AY2").Value + 1

    'TODO: Make this into a button

End Sub

Sub Reset()

    With Sheets("Current Generation")
        .Range("AY2").Value = 0
        .Range("C3:AP42").Value = 0
    End With
    'TODO: make into a button
End Sub


''TODO:
'Make buttons to start and stop the simulation
'Sample configurations
'randomize the cells
