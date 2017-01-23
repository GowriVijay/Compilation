'************************************************** US Macros Starts here *************************************************
'This macro is to transpose the SPSS output to fit the Excel template provided by Sophie

Sub US_inverse_macro()

Dim Data, Report As Worksheet
Dim j, N, Dr, Rp As Integer

' ***************************** Note Change the Sheet Name here *********************
Set Data = ThisWorkbook.Worksheets("US") ' Change the sheet name depending on the requirement

Data.Activate
Data.Range("Q3:DP400").ClearContents

Rp = 3

For j = 1 To 25 'For UK & US change 18 to 25
    'Copy & Transpose Paste module starts here
    Data.Range("A2:N105").Copy
    Data.Cells(Rp, 17).PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    'Copy & Transpose Paste ends here
    
    'Shifting cells up module starts here
      Data.Range("A1:N108").Delete Shift:=xlUp
    'Shifting cells up ends here
    
    Rp = Rp + 16
Next j

End Sub


'Macro for Creating Different Cuts that are required for reporting Authentication Survey
'This Macro can be used for Overall Numbers, US only Numbers, UK Only Numbers & DE Only Numbers
'However the Macro name might be misleading

Sub US_Report()
Dim Data, Report As Worksheet
Dim j, N, Dr, Rr As Integer


' ***************************** Note Change the Sheet Name here *********************
Set Data = ThisWorkbook.Worksheets("US")  ' Change the sheet name depending on the requirement
Set Report = ThisWorkbook.Worksheets("US Report") ' Change the sheet name depending on the requirement



R1 = 7
For Q = 1 To 25 'For UK & US change 18 to 25
    R2 = Cells(R1, 18).End(xlDown).Row
    
    Data.Activate
    Data.Range(Cells(R1, 18), Cells(R2, 18)).Select
    Selection.Copy
    Report.Activate
    Report.Cells(R1, 1).PasteSpecial xlPasteAll
    
    Data.Activate
    j = 2
    For X = 20 To 113
        N = X - 1
        Dr = R2 + 1
        Rr = R1 - 1
        
        
        Data.Range(Cells(R1, X), Cells(R2, X)).Select
        Selection.Copy
        Report.Activate
        Report.Cells(R1, j).PasteSpecial xlPasteAll
        Data.Activate
        Data.Cells(Dr, N).Copy
        Report.Activate
        Report.Cells(Rr, j).PasteSpecial xlPasteAll
        
        X = X + 2
        j = j + 1
        N = X - 1
        Data.Activate
        Data.Range(Cells(R1, X), Cells(R2, X)).Select
        Selection.Copy
        Report.Activate
        Report.Cells(R1, j).PasteSpecial xlPasteAll
        Data.Activate
        Data.Cells(Dr, N).Copy
        Report.Activate
        Report.Cells(Rr, j).PasteSpecial xlPasteAll
        
        X = X + 3
        j = j + 1
        Data.Activate
    Next X
Data.Cells(R1, 17).Font.Color = vbGreen

R1 = R1 + 16
Next Q
Report.Activate
Report.Range("A2").Select
End Sub

' Copying Data in Sophie's template

Sub template_paste_US()
Dim Data, Report As Worksheet
Dim Tmplt As Workbook

Set Report = ThisWorkbook.Worksheets("US Report") ' Change the sheet name depending on the requirement
Set Tmplt = Workbooks("Authentication buyer survey result analytics SH.xlsx")

'Q1
Report.Range("B6:AG11").Copy
Tmplt.Activate
Tmplt.Worksheets("Q1").Range("H22").PasteSpecial xlPasteValues

'Q2
Report.Range("B22:AG25").Copy
Tmplt.Activate
Tmplt.Worksheets("Q2").Range("H20").PasteSpecial xlPasteValues

'Q3 Base
Report.Activate
Report.Range("B38:AG38").Copy
Tmplt.Activate
Tmplt.Worksheets("Q3").Range("H24").PasteSpecial xlPasteValues

    'Q3a eBay Money Back Guarantee
    Report.Range("B40:AG40").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("H25").PasteSpecial xlPasteValues
    
    'Q3b Trust in the seller
    Report.Range("B56:AG56").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("H26").PasteSpecial xlPasteValues

    'Q3c Detailed item description
    Report.Range("B72:AG72").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("H27").PasteSpecial xlPasteValues

    'Q3d High quality images that show item details
    Report.Range("B88:AG88").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("H28").PasteSpecial xlPasteValues

    'Q3e Returns policy enabled me to return the item if I didn’t like it
    Report.Range("B104:AG104").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("H29").PasteSpecial xlPasteValues

    'Q3f Listing stated the item was authentic
    Report.Range("B120:AG120").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("H30").PasteSpecial xlPasteValues

    'Q3g Others (Please Specify)
    Report.Range("B136:AG136").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("H31").PasteSpecial xlPasteValues

'Q4a Base
Report.Activate
Report.Range("B150:AG150").Copy
Tmplt.Activate
Tmplt.Worksheets("Q4a").Range("H22").PasteSpecial xlPasteValues

    'Q4 Extremely Likely -  5
    Report.Range("B155:AG155").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4a").Range("H23").PasteSpecial xlPasteValues

    'Q4                     4
    Report.Range("B154:AG154").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4a").Range("H24").PasteSpecial xlPasteValues

    'Q4                     3
    Report.Range("B153:AG153").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4a").Range("H25").PasteSpecial xlPasteValues

    'Q4                     2
    Report.Range("B152:AG152").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4a").Range("H26").PasteSpecial xlPasteValues

    'Q4 Extremely Unlikely - 1
    Report.Range("B151:AG151").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4a").Range("H27").PasteSpecial xlPasteValues


'Q4b Base
Report.Activate
Report.Range("B166:AG166").Copy
Tmplt.Activate
Tmplt.Worksheets("Q4b").Range("H22").PasteSpecial xlPasteValues

    'Q4 Extremely Likely -  5
    Report.Range("B171:AG171").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4b").Range("H23").PasteSpecial xlPasteValues

    'Q4                     4
    Report.Range("B170:AG170").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4b").Range("H24").PasteSpecial xlPasteValues

    'Q4                     3
    Report.Range("B169:AG169").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4b").Range("H25").PasteSpecial xlPasteValues

    'Q4                     2
    Report.Range("B168:AG168").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4b").Range("H26").PasteSpecial xlPasteValues

    'Q4 Extremely Unlikely - 1
    Report.Range("B167:AG167").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4b").Range("H27").PasteSpecial xlPasteValues


'Q5
Report.Activate
Report.Range("B182:AG191").Copy
Tmplt.Activate
Tmplt.Worksheets("Q5").Range("H26").PasteSpecial xlPasteValues

'Q6
Report.Activate
Report.Range("B198:AG198").Copy
Tmplt.Activate
Tmplt.Worksheets("Q6").Range("H22").PasteSpecial xlPasteValues

    'Q6 Extremely Likely -  5
    Report.Range("B203:AG203").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q6").Range("H23").PasteSpecial xlPasteValues

    'Q6                     4
    Report.Range("B202:AG202").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q6").Range("H24").PasteSpecial xlPasteValues

    'Q6                     3
    Report.Range("B201:AG201").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q6").Range("H25").PasteSpecial xlPasteValues

    'Q6                     2
    Report.Range("B200:AG200").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q6").Range("H26").PasteSpecial xlPasteValues

    'Q6 Extremely Unlikely - 1
    Report.Range("B199:AG199").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q6").Range("H27").PasteSpecial xlPasteValues
    
'Q7
Report.Activate
Report.Range("B214:AG214").Copy
Tmplt.Activate
Tmplt.Worksheets("Q7").Range("H22").PasteSpecial xlPasteValues

    'Q7 Extremely Likely -  5
    Report.Range("B219:AG219").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q7").Range("H23").PasteSpecial xlPasteValues

    'Q7                     4
    Report.Range("B218:AG218").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q7").Range("H24").PasteSpecial xlPasteValues

    'Q7                     3
    Report.Range("B217:AG217").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q7").Range("H25").PasteSpecial xlPasteValues

    'Q7                     2
    Report.Range("B216:AG216").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q7").Range("H26").PasteSpecial xlPasteValues

    'Q7 Extremely Unlikely - 1
    Report.Range("B215:AG215").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q7").Range("H27").PasteSpecial xlPasteValues

'Q8
Report.Activate
Report.Range("B230:AG235").Copy
Tmplt.Activate
Tmplt.Worksheets("Q8").Range("H22").PasteSpecial xlPasteValues


'Q9 Base
Report.Activate
Report.Range("B246:AG246").Copy
Tmplt.Activate
Tmplt.Worksheets("Q9").Range("H21").PasteSpecial xlPasteValues

    'Q9a eBay brand / product experts
    Report.Range("B248:AG248").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q9").Range("H22").PasteSpecial xlPasteValues
    
    'Q9b Professional brand / product experts
    Report.Range("B264:AG264").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q9").Range("H23").PasteSpecial xlPasteValues

    'Q9c eBay Professional sellers who have expertise
    Report.Range("B280:AG280").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q9").Range("H24").PasteSpecial xlPasteValues

    'Q9d None of the above
    Report.Range("B296:AG296").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q9").Range("H25").PasteSpecial xlPasteValues
    
'Q10a
Report.Activate
Report.Range("B310:AG317").Copy
Tmplt.Activate
Tmplt.Worksheets("Q10a").Range("I24").PasteSpecial xlPasteValues

'Q10b
Report.Activate
Report.Range("B326:AG333").Copy
Tmplt.Activate
Tmplt.Worksheets("Q10b").Range("I24").PasteSpecial xlPasteValues

'Q10c
Report.Activate
Report.Range("B342:AG349").Copy
Tmplt.Activate
Tmplt.Worksheets("Q10c").Range("I24").PasteSpecial xlPasteValues

'Q10d
Report.Activate
Report.Range("B358:AG365").Copy
Tmplt.Activate
Tmplt.Worksheets("Q10d").Range("I24").PasteSpecial xlPasteValues

'Q10e
Report.Activate
Report.Range("B374:AG381").Copy
Tmplt.Activate
Tmplt.Worksheets("Q10e").Range("I24").PasteSpecial xlPasteValues

'Q10f
Report.Activate
Report.Range("B390:AG397").Copy
Tmplt.Activate
Tmplt.Worksheets("Q10f").Range("I24").PasteSpecial xlPasteValues
End Sub

' ************************************************** UK Macros Starts here *************************************************
'This macro is to transpose the SPSS output to fit the Excel template provided by Sophie

Sub UK_inverse_macro()

Dim Data, Report As Worksheet
Dim j, N, Dr, Rp As Integer

' ***************************** Note Change the Sheet Name here *********************
Set Data = ThisWorkbook.Worksheets("UK") ' Change the sheet name depending on the requirement

Data.Activate
Data.Range("Q3:DP400").ClearContents

Rp = 3

For j = 1 To 25 'For UK & US change 18 to 25
    'Copy & Transpose Paste module starts here
    Data.Range("A2:N105").Copy
    Data.Cells(Rp, 17).PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    'Copy & Transpose Paste ends here
    
    'Shifting cells up module starts here
      Data.Range("A1:N108").Delete Shift:=xlUp
    'Shifting cells up ends here
    
    Rp = Rp + 16
Next j

End Sub

'Macro for Creating Different Cuts that are required for reporting Authentication Survey
'This Macro can be used for Overall Numbers, US only Numbers, UK Only Numbers & DE Only Numbers
'However the Macro name might be misleading

Sub UK_Report()
Dim Data, Report As Worksheet
Dim j, N, Dr, Rr As Integer


' ***************************** Note Change the Sheet Name here *********************
Set Data = ThisWorkbook.Worksheets("UK")  ' Change the sheet name depending on the requirement
Set Report = ThisWorkbook.Worksheets("UK Report") ' Change the sheet name depending on the requirement



R1 = 7
For Q = 1 To 25 'For UK & US change 18 to 25
    R2 = Cells(R1, 18).End(xlDown).Row
    
    Data.Activate
    Data.Range(Cells(R1, 18), Cells(R2, 18)).Select
    Selection.Copy
    Report.Activate
    Report.Cells(R1, 1).PasteSpecial xlPasteAll
    
    Data.Activate
    j = 2
    For X = 20 To 113
        N = X - 1
        Dr = R2 + 1
        Rr = R1 - 1
        
        
        Data.Range(Cells(R1, X), Cells(R2, X)).Select
        Selection.Copy
        Report.Activate
        Report.Cells(R1, j).PasteSpecial xlPasteAll
        Data.Activate
        Data.Cells(Dr, N).Copy
        Report.Activate
        Report.Cells(Rr, j).PasteSpecial xlPasteAll
        
        X = X + 2
        j = j + 1
        N = X - 1
        Data.Activate
        Data.Range(Cells(R1, X), Cells(R2, X)).Select
        Selection.Copy
        Report.Activate
        Report.Cells(R1, j).PasteSpecial xlPasteAll
        Data.Activate
        Data.Cells(Dr, N).Copy
        Report.Activate
        Report.Cells(Rr, j).PasteSpecial xlPasteAll
        
        X = X + 3
        j = j + 1
        Data.Activate
    Next X
Data.Cells(R1, 17).Font.Color = vbGreen

R1 = R1 + 16
Next Q
Report.Activate
Report.Range("A2").Select
End Sub

' Copying the data into Sophie's template

Sub template_paste_UK()
Dim Data, Report As Worksheet
Dim Tmplt As Workbook

Set Report = ThisWorkbook.Worksheets("UK Report") ' Change the sheet name depending on the requirement
Set Tmplt = Workbooks("Authentication buyer survey result analytics SH.xlsx")

'Q1
Report.Range("B6:AG11").Copy
Tmplt.Activate
Tmplt.Worksheets("Q1").Range("H34").PasteSpecial xlPasteValues

'Q2
Report.Range("B22:AG25").Copy
Tmplt.Activate
Tmplt.Worksheets("Q2").Range("H30").PasteSpecial xlPasteValues

'Q3 Base
Report.Activate
Report.Range("B38:AG38").Copy
Tmplt.Activate
Tmplt.Worksheets("Q3").Range("H38").PasteSpecial xlPasteValues

    'Q3a eBay Money Back Guarantee
    Report.Range("B40:AG40").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("H39").PasteSpecial xlPasteValues
    
    'Q3b Trust in the seller
    Report.Range("B56:AG56").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("H40").PasteSpecial xlPasteValues

    'Q3c Detailed item description
    Report.Range("B72:AG72").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("H41").PasteSpecial xlPasteValues

    'Q3d High quality images that show item details
    Report.Range("B88:AG88").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("H42").PasteSpecial xlPasteValues

    'Q3e Returns policy enabled me to return the item if I didn’t like it
    Report.Range("B104:AG104").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("H43").PasteSpecial xlPasteValues

    'Q3f Listing stated the item was authentic
    Report.Range("B120:AG120").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("H44").PasteSpecial xlPasteValues

    'Q3g Others (Please Specify)
    Report.Range("B136:AG136").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("H45").PasteSpecial xlPasteValues

'Q4a Base
Report.Activate
Report.Range("B150:AG150").Copy
Tmplt.Activate
Tmplt.Worksheets("Q4a").Range("H34").PasteSpecial xlPasteValues

    'Q4 Extremely Likely -  5
    Report.Range("B155:AG155").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4a").Range("H35").PasteSpecial xlPasteValues

    'Q4                     4
    Report.Range("B154:AG154").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4a").Range("H36").PasteSpecial xlPasteValues

    'Q4                     3
    Report.Range("B153:AG153").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4a").Range("H37").PasteSpecial xlPasteValues

    'Q4                     2
    Report.Range("B152:AG152").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4a").Range("H38").PasteSpecial xlPasteValues

    'Q4 Extremely Unlikely - 1
    Report.Range("B151:AG151").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4a").Range("H39").PasteSpecial xlPasteValues


'Q4b Base
Report.Activate
Report.Range("B166:AG166").Copy
Tmplt.Activate
Tmplt.Worksheets("Q4b").Range("H34").PasteSpecial xlPasteValues

    'Q4 Extremely Likely -  5
    Report.Range("B171:AG171").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4b").Range("H35").PasteSpecial xlPasteValues

    'Q4                     4
    Report.Range("B170:AG170").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4b").Range("H36").PasteSpecial xlPasteValues

    'Q4                     3
    Report.Range("B169:AG169").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4b").Range("H37").PasteSpecial xlPasteValues

    'Q4                     2
    Report.Range("B168:AG168").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4b").Range("H38").PasteSpecial xlPasteValues

    'Q4 Extremely Unlikely - 1
    Report.Range("B167:AG167").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4b").Range("H39").PasteSpecial xlPasteValues


'Q5
Report.Activate
Report.Range("B182:AG191").Copy
Tmplt.Activate
Tmplt.Worksheets("Q5").Range("H42").PasteSpecial xlPasteValues

'Q6
Report.Activate
Report.Range("B198:AG198").Copy
Tmplt.Activate
Tmplt.Worksheets("Q6").Range("H34").PasteSpecial xlPasteValues

    'Q6 Extremely Likely -  5
    Report.Range("B203:AG203").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q6").Range("H35").PasteSpecial xlPasteValues

    'Q6                     4
    Report.Range("B202:AG202").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q6").Range("H36").PasteSpecial xlPasteValues

    'Q6                     3
    Report.Range("B201:AG201").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q6").Range("H37").PasteSpecial xlPasteValues

    'Q6                     2
    Report.Range("B200:AG200").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q6").Range("H38").PasteSpecial xlPasteValues

    'Q6 Extremely Unlikely - 1
    Report.Range("B199:AG199").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q6").Range("H39").PasteSpecial xlPasteValues
    
'Q7
Report.Activate
Report.Range("B214:AG214").Copy
Tmplt.Activate
Tmplt.Worksheets("Q7").Range("H34").PasteSpecial xlPasteValues

    'Q7 Extremely Likely -  5
    Report.Range("B219:AG219").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q7").Range("H35").PasteSpecial xlPasteValues

    'Q7                     4
    Report.Range("B218:AG218").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q7").Range("H36").PasteSpecial xlPasteValues

    'Q7                     3
    Report.Range("B217:AG217").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q7").Range("H37").PasteSpecial xlPasteValues

    'Q7                     2
    Report.Range("B216:AG216").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q7").Range("H38").PasteSpecial xlPasteValues

    'Q7 Extremely Unlikely - 1
    Report.Range("B215:AG215").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q7").Range("H39").PasteSpecial xlPasteValues

'Q8
Report.Activate
Report.Range("B230:AG235").Copy
Tmplt.Activate
Tmplt.Worksheets("Q8").Range("H34").PasteSpecial xlPasteValues


'Q9 Base
Report.Activate
Report.Range("B246:AG246").Copy
Tmplt.Activate
Tmplt.Worksheets("Q9").Range("H32").PasteSpecial xlPasteValues

    'Q9a eBay brand / product experts
    Report.Range("B248:AG248").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q9").Range("H33").PasteSpecial xlPasteValues
    
    'Q9b Professional brand / product experts
    Report.Range("B264:AG264").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q9").Range("H34").PasteSpecial xlPasteValues

    'Q9c eBay Professional sellers who have expertise
    Report.Range("B280:AG280").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q9").Range("H35").PasteSpecial xlPasteValues

    'Q9d None of the above
    Report.Range("B296:AG296").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q9").Range("H36").PasteSpecial xlPasteValues
    
'Q10a
Report.Activate
Report.Range("B310:AG317").Copy
Tmplt.Activate
Tmplt.Worksheets("Q10a").Range("I38").PasteSpecial xlPasteValues

'Q10b
Report.Activate
Report.Range("B326:AG333").Copy
Tmplt.Activate
Tmplt.Worksheets("Q10b").Range("I38").PasteSpecial xlPasteValues

'Q10c
Report.Activate
Report.Range("B342:AG349").Copy
Tmplt.Activate
Tmplt.Worksheets("Q10c").Range("I38").PasteSpecial xlPasteValues

'Q10d
Report.Activate
Report.Range("B358:AG365").Copy
Tmplt.Activate
Tmplt.Worksheets("Q10d").Range("I38").PasteSpecial xlPasteValues

'Q10e
Report.Activate
Report.Range("B374:AG381").Copy
Tmplt.Activate
Tmplt.Worksheets("Q10e").Range("I38").PasteSpecial xlPasteValues

'Q10f
Report.Activate
Report.Range("B390:AG397").Copy
Tmplt.Activate
Tmplt.Worksheets("Q10f").Range("I38").PasteSpecial xlPasteValues
End Sub


'********************************************** DE Macro Starts Here *******************************
'This macro is to transpose the SPSS output to fit the Excel template provided by Sophie

Sub DE_inverse_macro()

Dim Data, Report As Worksheet
Dim j, N, Dr, Rp As Integer

' ***************************** Note Change the Sheet Name here *********************
Set Data = ThisWorkbook.Worksheets("DE") ' Change the sheet name depending on the requirement

Data.Activate
Data.Range("Q3:DP400").ClearContents

Rp = 3

For j = 1 To 25 'For UK & US change 18 to 25
    'Copy & Transpose Paste module starts here
    Data.Range("A2:N105").Copy
    Data.Cells(Rp, 17).PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    'Copy & Transpose Paste ends here
    
    'Shifting cells up module starts here
      Data.Range("A1:N108").Delete Shift:=xlUp
    'Shifting cells up ends here
    
    Rp = Rp + 16
Next j

End Sub


'********************************************** DE Macro for Q6 to Q9d Starts Here *******************************
'This macro is to transpose the SPSS output to fit the Excel template provided by Sophie

Sub DE_inverse_macro_Only_Q6_to_Q9()

Dim Data, Report As Worksheet
Dim j, N, Dr, Rp As Integer

' ***************************** Note Change the Sheet Name here *********************
Set Data = ThisWorkbook.Worksheets("DE for Questions 6 - 9") ' Change the sheet name depending on the requirement

Data.Activate
Data.Range("Q3:DP400").ClearContents

Rp = 3

For j = 1 To 7 'For UK & US change  to 25
    'Copy & Transpose Paste module starts here
    Data.Range("A2:N105").Copy
    Data.Cells(Rp, 17).PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    'Copy & Transpose Paste ends here
    
    'Shifting cells up module starts here
      Data.Range("A1:N108").Delete Shift:=xlUp
    'Shifting cells up ends here
    
    Rp = Rp + 16
Next j

End Sub


'Macro for Creating Different Cuts that are required for reporting Authentication Survey
'This Macro can be used for Overall Numbers, US only Numbers, UK Only Numbers & DE Only Numbers
'However the Macro name might be misleading

Sub DE_Report()
Dim Data, Report As Worksheet
Dim j, N, Dr, Rr As Integer


' ***************************** Note Change the Sheet Name here *********************
Set Data = ThisWorkbook.Worksheets("DE")  ' Change the sheet name depending on the requirement
Set Report = ThisWorkbook.Worksheets("DE Report") ' Change the sheet name depending on the requirement



R1 = 7
For Q = 1 To 25
    R2 = Cells(R1, 18).End(xlDown).Row
    
    Data.Activate
    Data.Range(Cells(R1, 18), Cells(R2, 18)).Select
    Selection.Copy
    Report.Activate
    Report.Cells(R1, 1).PasteSpecial xlPasteAll
    
    Data.Activate
    j = 2
    For X = 20 To 113
        N = X - 1
        Dr = R2 + 1
        Rr = R1 - 1
        
        
        Data.Range(Cells(R1, X), Cells(R2, X)).Select
        Selection.Copy
        Report.Activate
        Report.Cells(R1, j).PasteSpecial xlPasteAll
        Data.Activate
        Data.Cells(Dr, N).Copy
        Report.Activate
        Report.Cells(Rr, j).PasteSpecial xlPasteAll
        
        X = X + 2
        j = j + 1
        N = X - 1
        Data.Activate
        Data.Range(Cells(R1, X), Cells(R2, X)).Select
        Selection.Copy
        Report.Activate
        Report.Cells(R1, j).PasteSpecial xlPasteAll
        Data.Activate
        Data.Cells(Dr, N).Copy
        Report.Activate
        Report.Cells(Rr, j).PasteSpecial xlPasteAll
        
        X = X + 3
        j = j + 1
        Data.Activate
    Next X
Data.Cells(R1, 17).Font.Color = vbGreen

R1 = R1 + 16
Next Q
Report.Activate
Report.Range("A2").Select
End Sub

' Copying the data into Sophie's template

Sub template_paste_DE()
Dim Data, Report As Worksheet
Dim Tmplt As Workbook

Set Report = ThisWorkbook.Worksheets("DE Report") ' Change the sheet name depending on the requirement
Set Tmplt = Workbooks("Authentication buyer survey result analytics SH.xlsx")

'Q1
Report.Range("B6:AG11").Copy
Tmplt.Activate
Tmplt.Worksheets("Q1").Range("H46").PasteSpecial xlPasteValues

'Q2
Report.Range("B22:AG25").Copy
Tmplt.Activate
Tmplt.Worksheets("Q2").Range("H40").PasteSpecial xlPasteValues

'Q3 Base
Report.Activate
Report.Range("B38:AG38").Copy
Tmplt.Activate
Tmplt.Worksheets("Q3").Range("H52").PasteSpecial xlPasteValues

    'Q3a eBay Money Back Guarantee
    Report.Range("B40:AG40").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("H53").PasteSpecial xlPasteValues
    
    'Q3b Trust in the seller
    Report.Range("B56:AG56").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("H54").PasteSpecial xlPasteValues

    'Q3c Detailed item description
    Report.Range("B72:AG72").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("H55").PasteSpecial xlPasteValues

    'Q3d High quality images that show item details
    Report.Range("B88:AG88").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("H56").PasteSpecial xlPasteValues

    'Q3e Returns policy enabled me to return the item if I didn’t like it
    Report.Range("B104:AG104").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("H57").PasteSpecial xlPasteValues

    'Q3f Listing stated the item was authentic
    Report.Range("B120:AG120").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("H58").PasteSpecial xlPasteValues

    'Q3g Others (Please Specify)
    Report.Range("B136:AG136").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("H59").PasteSpecial xlPasteValues

'Q4a Base
Report.Activate
Report.Range("B150:AG150").Copy
Tmplt.Activate
Tmplt.Worksheets("Q4a").Range("H46").PasteSpecial xlPasteValues

    'Q4 Extremely Likely -  5
    Report.Range("B155:AG155").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4a").Range("H47").PasteSpecial xlPasteValues

    'Q4                     4
    Report.Range("B154:AG154").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4a").Range("H48").PasteSpecial xlPasteValues

    'Q4                     3
    Report.Range("B153:AG153").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4a").Range("H49").PasteSpecial xlPasteValues

    'Q4                     2
    Report.Range("B152:AG152").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4a").Range("H50").PasteSpecial xlPasteValues

    'Q4 Extremely Unlikely - 1
    Report.Range("B151:AG151").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4a").Range("H51").PasteSpecial xlPasteValues


'Q4b Base
Report.Activate
Report.Range("B166:AG166").Copy
Tmplt.Activate
Tmplt.Worksheets("Q4b").Range("H46").PasteSpecial xlPasteValues

    'Q4 Extremely Likely -  5
    Report.Range("B171:AG171").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4b").Range("H47").PasteSpecial xlPasteValues

    'Q4                     4
    Report.Range("B170:AG170").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4b").Range("H48").PasteSpecial xlPasteValues

    'Q4                     3
    Report.Range("B169:AG169").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4b").Range("H49").PasteSpecial xlPasteValues

    'Q4                     2
    Report.Range("B168:AG168").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4b").Range("H50").PasteSpecial xlPasteValues

    'Q4 Extremely Unlikely - 1
    Report.Range("B167:AG167").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4b").Range("H51").PasteSpecial xlPasteValues


'Q5
Report.Activate
Report.Range("B182:AG191").Copy
Tmplt.Activate
Tmplt.Worksheets("Q5").Range("H58").PasteSpecial xlPasteValues

'Q6
Report.Activate
Report.Range("B198:AG198").Copy
Tmplt.Activate
Tmplt.Worksheets("Q6").Range("H46").PasteSpecial xlPasteValues

    'Q6 Extremely Likely -  5
    Report.Range("B203:AG203").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q6").Range("H47").PasteSpecial xlPasteValues

    'Q6                     4
    Report.Range("B202:AG202").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q6").Range("H48").PasteSpecial xlPasteValues

    'Q6                     3
    Report.Range("B201:AG201").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q6").Range("H49").PasteSpecial xlPasteValues

    'Q6                     2
    Report.Range("B200:AG200").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q6").Range("H50").PasteSpecial xlPasteValues

    'Q6 Extremely Unlikely - 1
    Report.Range("B199:AG199").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q6").Range("H51").PasteSpecial xlPasteValues
    
'Q7
Report.Activate
Report.Range("B214:AG214").Copy
Tmplt.Activate
Tmplt.Worksheets("Q7").Range("H46").PasteSpecial xlPasteValues

    'Q7 Extremely Likely -  5
    Report.Range("B219:AG219").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q7").Range("H47").PasteSpecial xlPasteValues

    'Q7                     4
    Report.Range("B218:AG218").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q7").Range("H48").PasteSpecial xlPasteValues

    'Q7                     3
    Report.Range("B217:AG217").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q7").Range("H49").PasteSpecial xlPasteValues

    'Q7                     2
    Report.Range("B216:AG216").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q7").Range("H50").PasteSpecial xlPasteValues

    'Q7 Extremely Unlikely - 1
    Report.Range("B215:AG215").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q7").Range("H51").PasteSpecial xlPasteValues

'Q8
Report.Activate
Report.Range("B230:AG235").Copy
Tmplt.Activate
Tmplt.Worksheets("Q8").Range("H46").PasteSpecial xlPasteValues


'Q9 Base
Report.Activate
Report.Range("B246:AG246").Copy
Tmplt.Activate
Tmplt.Worksheets("Q9").Range("H43").PasteSpecial xlPasteValues

    'Q9a eBay brand / product experts
    Report.Range("B248:AG248").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q9").Range("H44").PasteSpecial xlPasteValues
    
    'Q9b Professional brand / product experts
    Report.Range("B264:AG264").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q9").Range("H45").PasteSpecial xlPasteValues

    'Q9c eBay Professional sellers who have expertise
    Report.Range("B280:AG280").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q9").Range("H46").PasteSpecial xlPasteValues

    'Q9d None of the above
    Report.Range("B296:AG296").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q9").Range("H47").PasteSpecial xlPasteValues
    
'Q10a
Report.Activate
Report.Range("B310:AG317").Copy
Tmplt.Activate
Tmplt.Worksheets("Q10a").Range("I52").PasteSpecial xlPasteValues

'Q10b
Report.Activate
Report.Range("B326:AG333").Copy
Tmplt.Activate
Tmplt.Worksheets("Q10b").Range("I52").PasteSpecial xlPasteValues

'Q10c
Report.Activate
Report.Range("B342:AG349").Copy
Tmplt.Activate
Tmplt.Worksheets("Q10c").Range("I52").PasteSpecial xlPasteValues

'Q10d
Report.Activate
Report.Range("B358:AG365").Copy
Tmplt.Activate
Tmplt.Worksheets("Q10d").Range("I52").PasteSpecial xlPasteValues

'Q10e
Report.Activate
Report.Range("B374:AG381").Copy
Tmplt.Activate
Tmplt.Worksheets("Q10e").Range("I52").PasteSpecial xlPasteValues

'Q10f
Report.Activate
Report.Range("B390:AG397").Copy
Tmplt.Activate
Tmplt.Worksheets("Q10f").Range("I52").PasteSpecial xlPasteValues
End Sub


'************************************************** All Macros Starts here *************************************************
'This macro is to transpose the SPSS output to fit the Excel template provided by Sophie

Sub All_inverse_macro()

Dim Data, Report As Worksheet
Dim j, N, Dr, Rp As Integer

' ***************************** Note Change the Sheet Name here *********************
Set Data = ThisWorkbook.Worksheets("All") ' Change the sheet name depending on the requirement

Data.Activate
Data.Range("Q3:DP400").ClearContents

Rp = 3

For j = 1 To 25 'For UK & US change 18 to 25
    'Copy & Transpose Paste module starts here
    Data.Range("A2:N105").Copy
    Data.Cells(Rp, 17).PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    'Copy & Transpose Paste ends here
    
    'Shifting cells up module starts here
      Data.Range("A1:N108").Delete Shift:=xlUp
    'Shifting cells up ends here
    
    Rp = Rp + 16
Next j

End Sub


'Macro for Creating Different Cuts that are required for reporting Authentication Survey
'This Macro can be used for Overall Numbers, US only Numbers, UK Only Numbers & DE Only Numbers
'However the Macro name might be misleading

Sub All_Report()
Dim Data, Report As Worksheet
Dim j, N, Dr, Rr As Integer


' ***************************** Note Change the Sheet Name here *********************
Set Data = ThisWorkbook.Worksheets("All")  ' Change the sheet name depending on the requirement
Set Report = ThisWorkbook.Worksheets("All Report") ' Change the sheet name depending on the requirement



R1 = 7
For Q = 1 To 25 'For UK & US change 18 to 25
    R2 = Cells(R1, 18).End(xlDown).Row
    
    Data.Activate
    Data.Range(Cells(R1, 18), Cells(R2, 18)).Select
    Selection.Copy
    Report.Activate
    Report.Cells(R1, 1).PasteSpecial xlPasteAll
    
    Data.Activate
    j = 2
    For X = 20 To 113
        N = X - 1
        Dr = R2 + 1
        Rr = R1 - 1
        
        
        Data.Range(Cells(R1, X), Cells(R2, X)).Select
        Selection.Copy
        Report.Activate
        Report.Cells(R1, j).PasteSpecial xlPasteAll
        Data.Activate
        Data.Cells(Dr, N).Copy
        Report.Activate
        Report.Cells(Rr, j).PasteSpecial xlPasteAll
        
        X = X + 2
        j = j + 1
        N = X - 1
        Data.Activate
        Data.Range(Cells(R1, X), Cells(R2, X)).Select
        Selection.Copy
        Report.Activate
        Report.Cells(R1, j).PasteSpecial xlPasteAll
        Data.Activate
        Data.Cells(Dr, N).Copy
        Report.Activate
        Report.Cells(Rr, j).PasteSpecial xlPasteAll
        
        X = X + 3
        j = j + 1
        Data.Activate
    Next X
Data.Cells(R1, 17).Font.Color = vbGreen

R1 = R1 + 16
Next Q
Report.Activate
Report.Range("A2").Select
End Sub

' Copying Data in Sophie's template

Sub template_paste_All()
Dim Data, Report As Worksheet
Dim Tmplt As Workbook

Set Report = ThisWorkbook.Worksheets("All Report") ' Change the sheet name depending on the requirement
Set Tmplt = Workbooks("Authentication buyer survey result analytics SH.xlsx")

'Q1
Report.Range("B6:AG11").Copy
Tmplt.Activate
Tmplt.Worksheets("Q1").Range("H10").PasteSpecial xlPasteValues

'Q2
Report.Range("B22:AG25").Copy
Tmplt.Activate
Tmplt.Worksheets("Q2").Range("H10").PasteSpecial xlPasteValues

'Q3 Base
Report.Activate
Report.Range("B38:AG38").Copy
Tmplt.Activate
Tmplt.Worksheets("Q3").Range("H10").PasteSpecial xlPasteValues

    'Q3a eBay Money Back Guarantee
    Report.Range("B40:AG40").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("H11").PasteSpecial xlPasteValues
    
    'Q3b Trust in the seller
    Report.Range("B56:AG56").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("H12").PasteSpecial xlPasteValues

    'Q3c Detailed item description
    Report.Range("B72:AG72").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("H13").PasteSpecial xlPasteValues

    'Q3d High quality images that show item details
    Report.Range("B88:AG88").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("H14").PasteSpecial xlPasteValues

    'Q3e Returns policy enabled me to return the item if I didn’t like it
    Report.Range("B104:AG104").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("H15").PasteSpecial xlPasteValues

    'Q3f Listing stated the item was authentic
    Report.Range("B120:AG120").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("H16").PasteSpecial xlPasteValues

    'Q3g Others (Please Specify)
    Report.Range("B136:AG136").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("H17").PasteSpecial xlPasteValues

'Q4a Base
Report.Activate
Report.Range("B150:AG150").Copy
Tmplt.Activate
Tmplt.Worksheets("Q4a").Range("H10").PasteSpecial xlPasteValues

    'Q4 Extremely Likely -  5
    Report.Range("B155:AG155").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4a").Range("H11").PasteSpecial xlPasteValues

    'Q4                     4
    Report.Range("B154:AG154").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4a").Range("H12").PasteSpecial xlPasteValues

    'Q4                     3
    Report.Range("B153:AG153").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4a").Range("H13").PasteSpecial xlPasteValues

    'Q4                     2
    Report.Range("B152:AG152").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4a").Range("H14").PasteSpecial xlPasteValues

    'Q4 Extremely Unlikely - 1
    Report.Range("B151:AG151").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4a").Range("H15").PasteSpecial xlPasteValues


'Q4b Base
Report.Activate
Report.Range("B166:AG166").Copy
Tmplt.Activate
Tmplt.Worksheets("Q4b").Range("H10").PasteSpecial xlPasteValues

    'Q4 Extremely Likely -  5
    Report.Range("B171:AG171").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4b").Range("H11").PasteSpecial xlPasteValues

    'Q4                     4
    Report.Range("B170:AG170").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4b").Range("H12").PasteSpecial xlPasteValues

    'Q4                     3
    Report.Range("B169:AG169").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4b").Range("H13").PasteSpecial xlPasteValues

    'Q4                     2
    Report.Range("B168:AG168").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4b").Range("H14").PasteSpecial xlPasteValues

    'Q4 Extremely Unlikely - 1
    Report.Range("B167:AG167").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4b").Range("H15").PasteSpecial xlPasteValues


'Q5
Report.Activate
Report.Range("B182:AG191").Copy
Tmplt.Activate
Tmplt.Worksheets("Q5").Range("H10").PasteSpecial xlPasteValues

'Q6
Report.Activate
Report.Range("B198:AG198").Copy
Tmplt.Activate
Tmplt.Worksheets("Q6").Range("H10").PasteSpecial xlPasteValues

    'Q6 Extremely Likely -  5
    Report.Range("B203:AG203").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q6").Range("H11").PasteSpecial xlPasteValues

    'Q6                     4
    Report.Range("B202:AG202").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q6").Range("H12").PasteSpecial xlPasteValues

    'Q6                     3
    Report.Range("B201:AG201").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q6").Range("H13").PasteSpecial xlPasteValues

    'Q6                     2
    Report.Range("B200:AG200").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q6").Range("H14").PasteSpecial xlPasteValues

    'Q6 Extremely Unlikely - 1
    Report.Range("B199:AG199").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q6").Range("H15").PasteSpecial xlPasteValues
    
'Q7
Report.Activate
Report.Range("B214:AG214").Copy
Tmplt.Activate
Tmplt.Worksheets("Q7").Range("H10").PasteSpecial xlPasteValues

    'Q7 Extremely Likely -  5
    Report.Range("B219:AG219").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q7").Range("H11").PasteSpecial xlPasteValues

    'Q7                     4
    Report.Range("B218:AG218").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q7").Range("H12").PasteSpecial xlPasteValues

    'Q7                     3
    Report.Range("B217:AG217").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q7").Range("H13").PasteSpecial xlPasteValues

    'Q7                     2
    Report.Range("B216:AG216").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q7").Range("H14").PasteSpecial xlPasteValues

    'Q7 Extremely Unlikely - 1
    Report.Range("B215:AG215").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q7").Range("H15").PasteSpecial xlPasteValues

'Q8
Report.Activate
Report.Range("B230:AG235").Copy
Tmplt.Activate
Tmplt.Worksheets("Q8").Range("H10").PasteSpecial xlPasteValues


'Q9 Base
Report.Activate
Report.Range("B246:AG246").Copy
Tmplt.Activate
Tmplt.Worksheets("Q9").Range("H10").PasteSpecial xlPasteValues

    'Q9a eBay brand / product experts
    Report.Range("B248:AG248").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q9").Range("H11").PasteSpecial xlPasteValues
    
    'Q9b Professional brand / product experts
    Report.Range("B264:AG264").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q9").Range("H12").PasteSpecial xlPasteValues

    'Q9c eBay Professional sellers who have expertise
    Report.Range("B280:AG280").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q9").Range("H13").PasteSpecial xlPasteValues

    'Q9d None of the above
    Report.Range("B296:AG296").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q9").Range("H14").PasteSpecial xlPasteValues
    
'Q10a
Report.Activate
Report.Range("B310:AG317").Copy
Tmplt.Activate
Tmplt.Worksheets("Q10a").Range("I10").PasteSpecial xlPasteValues

'Q10b
Report.Activate
Report.Range("B326:AG333").Copy
Tmplt.Activate
Tmplt.Worksheets("Q10b").Range("I10").PasteSpecial xlPasteValues

'Q10c
Report.Activate
Report.Range("B342:AG349").Copy
Tmplt.Activate
Tmplt.Worksheets("Q10c").Range("I10").PasteSpecial xlPasteValues

'Q10d
Report.Activate
Report.Range("B358:AG365").Copy
Tmplt.Activate
Tmplt.Worksheets("Q10d").Range("I10").PasteSpecial xlPasteValues

'Q10e
Report.Activate
Report.Range("B374:AG381").Copy
Tmplt.Activate
Tmplt.Worksheets("Q10e").Range("I10").PasteSpecial xlPasteValues

'Q10f
Report.Activate
Report.Range("B390:AG397").Copy
Tmplt.Activate
Tmplt.Worksheets("Q10f").Range("I10").PasteSpecial xlPasteValues
End Sub
