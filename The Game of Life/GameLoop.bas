Attribute VB_Name = "GameLoop"
Option Explicit

Sub GameLoop()
    
    ' Keep the game running until we hit a specified generation
    While Range("AY2").Value < Range("BF2").Value
        
        'Copy the successor generation to the clipboard and then
        '   paste it over the current generation
        Sheets("Successor Generation").Range("C3:AP42").Copy
        Sheets("Current Generation").Range("C3").PasteSpecial Paste:=xlPasteValues, _
            Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
        ' Update the generation value
        Range("AY2").Value = Range("AY2").Value + 1
    
    Wend
           
End Sub


''TODO:
'Make buttons to start and stop the simulation
'Make a button to reset it
'See if I can set a speed with pauses
'change formulas to wrap around the field
'Sample configurations
'randomize the cells
'set a button to run through generations manually
