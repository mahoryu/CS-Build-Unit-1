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
    Dim current() As Variant
    Dim temp(0 To 39, 0 To 39) As Variant
    Dim successor(1 To 40, 1 To 40) As Variant
    Dim Rng As Range
        
    Set Rng = Sheets("Current Generation").Range("C3:AP42")
    current = Rng
    
    ' move the array to a 0 based index array because of a limitation with VBA
    ' (Arrays taken from the sheet are index base 1 and can't be quickly changed
    '  so I am using a for loop to move the data over)
    Dim i As Integer
    For i = LBound(current, 1) To UBound(current, 1)
        Dim j As Integer
        For j = LBound(current, 2) To UBound(current, 2)
            temp(i - 1, j - 1) = current(i, j)
        Next
    Next
    
    ' loop through the base-0 array
    Dim N As Integer
    N = UBound(temp, 1) + 1
    For i = LBound(temp, 1) To UBound(temp, 1)
        For j = LBound(temp, 2) To UBound(temp, 2)
                
            Dim total As Integer
            
            ' Check for the conditionals that set the rules
            total = (temp(i, (j - 1 + N) Mod N) + _
                     temp(i, (j + 1 + N) Mod N) + _
                     temp((i - 1 + N) Mod N, j) + _
                     temp((i + 1 + N) Mod N, j) + _
                     temp((i - 1 + N) Mod N, (j - 1 + N) Mod N) + _
                     temp((i - 1 + N) Mod N, (j + 1 + N) Mod N) + _
                     temp((i + 1 + N) Mod N, (j + 1 + N) Mod N) + _
                     temp((i + 1 + N) Mod N, (j - 1 + N) Mod N))
                     
            ' Apply the logic and put it in the new array
            '   (base-1 so that it will paste properly to the sheet)
            If temp(i, j) = 1 Then
                If total = 2 Or total = 3 Then
                    successor(i + 1, j + 1) = 1
                Else
                    successor(i + 1, j + 1) = 0
                End If
            Else
                If total = 3 Then
                    successor(i + 1, j + 1) = 1
                Else
                    successor(i + 1, j + 1) = 0
                End If
            End If
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
