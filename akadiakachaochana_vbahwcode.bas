Attribute VB_Name = "Module1"
Sub Runallwks() 'I found this code to run my code on multiple sheets from this website, https://www.extendoffice.com/documents/excel/5333-excel-run-macro-multiple-sheets.html
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In worksheets
        xSh.Select
        Call vba_hw
    Next
    Application.ScreenUpdating = True
End Sub


Sub vba_hw()

Dim firstticker, nextticker As String
Dim lastrow, volumecounter, firstvolume, nextvolume As Long
Dim openprice, closeprice, yrchange, percchange As Double
Dim outputrow As Integer

lastrow = Cells(Rows.Count, 1).End(xlUp).Row
'Added code below to sort data by ticker name first, then by date, to ensure all data is sorted before applying sub.
Range("A:G").Sort key1:=Range("A1"), order1:=xlAscending, key2:=Range("B1"), order2:=xlAscending, Header:=xlYes

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(1, 14).Value = "Bonus questions"
Cells(1, 15).Value = "Ticker"
Cells(1, 16).Value = "Value"
Cells(2, 14).Value = "Greatest % Increase"
Cells(3, 14).Value = "Greatest % Decrease"
Cells(4, 14).Value = "Greatest Total Volume"

outputrow = 2
firstticker = Cells(2, 1).Value
    Cells(outputrow, 9) = firstticker
openprice = Cells(2, 3).Value
volumecounter = Cells(2, 7).Value

For I = 0 To lastrow - 2 'Set this range so that with the code below,will loop through all the rows with ticker values (vs. header and blank rows)

    If Cells(3 + I, 1).Value <> Cells(2 + I, 1).Value Then
        outputrow = outputrow + 1
        closeprice = Cells(2 + I, 6).Value
        nextticker = Cells(3 + I, 1).Value
            Cells(outputrow, 9).Value = nextticker
        yrchange = (closeprice - openprice)
            Cells(outputrow - 1, 10).Value = yrchange
                If yrchange < 0 Then Cells(outputrow - 1, 10).Interior.ColorIndex = 3
                If yrchange > 0 Then Cells(outputrow - 1, 10).Interior.ColorIndex = 4
        If openprice <> 0 Then percchange = ((closeprice - openprice) / openprice) 'needed to add this If statement to debug error of dividing by 0
            Cells(outputrow - 1, 11).Value = Format(percchange, "Percent")
        volumecounter = Cells(2 + I, 7).Value + volumecounter
            Cells(outputrow - 1, 12).Value = volumecounter
        volumecounter = Cells(3 + I, 7).Value
        openprice = Cells(3 + I, 3).Value
            
    Else
        volumecounter = Cells(2 + I, 7).Value + volumecounter
                
   End If
   
Next I

Dim lastrowoutput As Integer
Dim highestvol As Double
Dim highestticker As String

lastrowoutput = Cells(Rows.Count, 12).End(xlUp).Row
highestvol = Cells(2, 12).Value
highestvolticker = Cells(2, 9).Value

For I = 2 To lastrowoutput
    If Cells(I + 1, 12).Value > highestvol Then
        highestvol = Cells(I + 1, 12).Value
        highestvolticker = Cells(I + 1, 9).Value

    End If
    
    Next I
    
        Cells(4, 16).Value = highestvol
        Cells(4, 15).Value = highestvolticker

Dim highestpc As Double
Dim highestpcticker As String

lastrowoutput = Cells(Rows.Count, 12).End(xlUp).Row
highestpc = Cells(2, 11).Value
highestpcticker = Cells(2, 9).Value

For I = 2 To lastrowoutput
    If Cells(I + 1, 11).Value > highestpc Then
        highestpc = Cells(I + 1, 11).Value
        highestpcticker = Cells(I + 1, 9).Value

    End If
    
    Next I
    
        Cells(2, 16).Value = Format(highestpc, "Percent")
        Cells(2, 15).Value = highestpcticker

Dim lowestpc As Double
Dim lowestpcticker As String

lastrowoutput = Cells(Rows.Count, 12).End(xlUp).Row
lowestpc = Cells(2, 11).Value
lowestpcticker = Cells(2, 9).Value

For I = 2 To lastrowoutput
    If Cells(I + 1, 11).Value < lowestpc Then
        lowestpc = Cells(I + 1, 11).Value
        lowestpcticker = Cells(I + 1, 9).Value

    End If
    
    Next I
    
        Cells(3, 16).Value = Format(lowestpc, "Percent")
        Cells(3, 15).Value = lowestpcticker
End Sub
