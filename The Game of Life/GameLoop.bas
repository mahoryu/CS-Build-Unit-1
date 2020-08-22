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
        
        Call GameStep
        
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

    'store the data in an array and asign it from successor to current
        Dim data() As Variant
        Dim Rng As Range
        
        Set Rng = Sheets("Successor Generation").Range("C3:AP42")
        data = Rng
        Sheets("Current Generation").Range("C3:AP42") = data
    
    ' Update the generation value
    Range("AY2").Value = Range("AY2").Value + 1

End Sub

Sub GameStep_v2()

    'store the data in an array and asign it from successor to current
    Dim current() As Variant
    Dim successor() As Variant
    Dim Rng As Range
        
    Set Rng = Sheets("Current Generation").Range("C3:AP42")
    current = Rng
    successor = current
        
    Debug.Print "Start"
    Dim i As Integer
    Dim N As Integer
    N = UBound(current, 1)
    For i = LBound(current, 1) To UBound(current, 1)
        Dim j As Integer
        For j = LBound(current, 2) To UBound(current, 2)
                
            Dim total As Integer
            
            total = (current(i, (j - 2 + N) Mod N + 1) + _
                     current(i, (j + 1 + N) Mod N) + _
                     current((i - 2 + N) Mod N + 1, j) + _
                     current((i + 1 + N) Mod N, j) + _
                     current((i - 2 + N) Mod N + 1, (j - 2 + N) Mod N + 1) + _
                     current((i - 2 + N) Mod N + 1, (j + 1 + N) Mod N) + _
                     current((i + 1 + N) Mod N, (j + 1 + N) Mod N) + _
                     current((i + 1 + N) Mod N, (j - 2 + N) Mod N + 1))
            Next
        Next
        
        Sheets("Current Generation").Range("C3:AP42") = successor
    
    ' Update the generation value
    Range("AY2").Value = Range("AY2").Value + 1

End Sub

Sub Reset()
    ' Resets the Generation and sets all values to dead
    With Sheets("Current Generation")
        .Range("AY2").Value = 0
        .Range("C3:AP42").Value = 0
    End With
End Sub


''TODO:
'Sample configurations
'randomize the cells


Sub testing()
    Dim N As Integer
    N = 40
    Debug.Print (1 - 2 + N) Mod N + 1
    Debug.Print (39 + 1) Mod (N)
End Sub
