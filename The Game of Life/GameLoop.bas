Attribute VB_Name = "GameLoop"
Option Explicit

Sub GameLoop()
        
    'Copy the successor generation to the clipboard and then
    '   paste it over the current generation
    Sheets("Successor Generation").Range("C3:AP42").Copy
    Sheets("Current Generation").Range("C3").PasteSpecial Paste:=xlPasteValues, _
        Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    ' Update the generation value
    Range("AY2").Value = Range("AY2").Value + 1
    
        
End Sub
