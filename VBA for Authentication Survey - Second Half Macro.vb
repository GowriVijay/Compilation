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
    Data.Range("A2:N21").Copy
    Data.Cells(Rp, 17).PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    'Copy & Transpose Paste ends here
    
    'Shifting cells up module starts here
      Data.Range("A1:N24").Delete Shift:=xlUp
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
    Data.Activate
    R2 = Cells(R1, 18).End(xlDown).Row
    
  
    Data.Range(Cells(R1, 18), Cells(R2, 18)).Select
    Selection.Copy
    Report.Activate
    Report.Cells(R1, 1).PasteSpecial xlPasteAll
    
    Data.Activate
    j = 3
    For X = 20 To 29
        N = X - 1 'For calculating the sample size for each Question
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
Data.Range(Cells(R1, 36), Cells(R2, 36)).Select
Selection.Copy
Report.Activate
Report.Cells(R1, 2).PasteSpecial xlPasteAll
Report.Cells(R1 - 1, 2).FormulaR1C1 = "=SUM(RC[1]:RC[4])"

R1 = R1 + 16
Data.Cells(R1, 17).Font.Color = vbGreen

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
Report.Range("B6:F11").Copy
Tmplt.Activate
Tmplt.Worksheets("Q1").Range("C22").PasteSpecial xlPasteValues

'Q2
Report.Range("B22:F25").Copy
Tmplt.Activate
Tmplt.Worksheets("Q2").Range("C20").PasteSpecial xlPasteValues

'Q3 Base
Report.Activate
Report.Range("B38:F38").Copy
Tmplt.Activate
Tmplt.Worksheets("Q3").Range("C24").PasteSpecial xlPasteValues

    'Q3a eBay Money Back Guarantee
    Report.Range("B40:F40").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("C25").PasteSpecial xlPasteValues
    
    'Q3b Trust in the seller
    Report.Range("B56:F56").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("C26").PasteSpecial xlPasteValues

    'Q3c Detailed item description
    Report.Range("B72:F72").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("C27").PasteSpecial xlPasteValues

    'Q3d High quality images that show item details
    Report.Range("B88:F88").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("C28").PasteSpecial xlPasteValues

    'Q3e Returns policy enabled me to return the item if I didn’t like it
    Report.Range("B104:F104").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("C29").PasteSpecial xlPasteValues

    'Q3f Listing stated the item was authentic
    Report.Range("B120:F120").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("C30").PasteSpecial xlPasteValues

    'Q3g Others (Please Specify)
    Report.Range("B136:F136").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("C31").PasteSpecial xlPasteValues

'Q4a Base
Report.Activate
Report.Range("B150:F150").Copy
Tmplt.Activate
Tmplt.Worksheets("Q4a").Range("C22").PasteSpecial xlPasteValues

    'Q4 Extremely Likely -  5
    Report.Range("B155:F155").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4a").Range("C23").PasteSpecial xlPasteValues

    'Q4                     4
    Report.Range("B154:F154").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4a").Range("C24").PasteSpecial xlPasteValues

    'Q4                     3
    Report.Range("B153:F153").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4a").Range("C25").PasteSpecial xlPasteValues

    'Q4                     2
    Report.Range("B152:F152").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4a").Range("C26").PasteSpecial xlPasteValues

    'Q4 Extremely Unlikely - 1
    Report.Range("B151:F151").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4a").Range("C27").PasteSpecial xlPasteValues


'Q4b Base
Report.Activate
Report.Range("B166:F166").Copy
Tmplt.Activate
Tmplt.Worksheets("Q4b").Range("C22").PasteSpecial xlPasteValues

    'Q4 Extremely Likely -  5
    Report.Range("B171:F171").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4b").Range("C23").PasteSpecial xlPasteValues

    'Q4                     4
    Report.Range("B170:F170").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4b").Range("C24").PasteSpecial xlPasteValues

    'Q4                     3
    Report.Range("B169:F169").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4b").Range("C25").PasteSpecial xlPasteValues

    'Q4                     2
    Report.Range("B168:F168").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4b").Range("C26").PasteSpecial xlPasteValues

    'Q4 Extremely Unlikely - 1
    Report.Range("B167:F167").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4b").Range("C27").PasteSpecial xlPasteValues


'Q5
Report.Activate
Report.Range("B182:F191").Copy
Tmplt.Activate
Tmplt.Worksheets("Q5").Range("C26").PasteSpecial xlPasteValues

'Q6
Report.Activate
Report.Range("B198:F198").Copy
Tmplt.Activate
Tmplt.Worksheets("Q6").Range("C22").PasteSpecial xlPasteValues

    'Q6 Extremely Likely -  5
    Report.Range("B203:F203").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q6").Range("C23").PasteSpecial xlPasteValues

    'Q6                     4
    Report.Range("B202:F202").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q6").Range("C24").PasteSpecial xlPasteValues

    'Q6                     3
    Report.Range("B201:F201").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q6").Range("C25").PasteSpecial xlPasteValues

    'Q6                     2
    Report.Range("B200:F200").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q6").Range("C26").PasteSpecial xlPasteValues

    'Q6 Extremely Unlikely - 1
    Report.Range("B199:F199").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q6").Range("C27").PasteSpecial xlPasteValues
    
'Q7
Report.Activate
Report.Range("B214:F214").Copy
Tmplt.Activate
Tmplt.Worksheets("Q7").Range("C22").PasteSpecial xlPasteValues

    'Q7 Extremely Likely -  5
    Report.Range("B219:F219").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q7").Range("C23").PasteSpecial xlPasteValues

    'Q7                     4
    Report.Range("B218:F218").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q7").Range("C24").PasteSpecial xlPasteValues

    'Q7                     3
    Report.Range("B217:F217").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q7").Range("C25").PasteSpecial xlPasteValues

    'Q7                     2
    Report.Range("B216:F216").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q7").Range("C26").PasteSpecial xlPasteValues

    'Q7 Extremely Unlikely - 1
    Report.Range("B215:F215").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q7").Range("C27").PasteSpecial xlPasteValues

'Q8
Report.Activate
Report.Range("B230:F235").Copy
Tmplt.Activate
Tmplt.Worksheets("Q8").Range("C22").PasteSpecial xlPasteValues


'Q9 Base
Report.Activate
Report.Range("B246:F246").Copy
Tmplt.Activate
Tmplt.Worksheets("Q9").Range("C21").PasteSpecial xlPasteValues

    'Q9a eBay brand / product experts
    Report.Range("B248:F248").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q9").Range("C22").PasteSpecial xlPasteValues
    
    'Q9b Professional brand / product experts
    Report.Range("B264:F264").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q9").Range("C23").PasteSpecial xlPasteValues

    'Q9c eBay Professional sellers who have expertise
    Report.Range("B280:F280").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q9").Range("C24").PasteSpecial xlPasteValues

    'Q9d None of the above
    Report.Range("B296:F296").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q9").Range("C25").PasteSpecial xlPasteValues
    
'Q10a
Report.Activate
Report.Range("B310:F317").Copy
Tmplt.Activate
Tmplt.Worksheets("Q10a").Range("D24").PasteSpecial xlPasteValues

'Q10b
Report.Activate
Report.Range("B326:F333").Copy
Tmplt.Activate
Tmplt.Worksheets("Q10b").Range("D24").PasteSpecial xlPasteValues

'Q10c
Report.Activate
Report.Range("B342:F349").Copy
Tmplt.Activate
Tmplt.Worksheets("Q10c").Range("D24").PasteSpecial xlPasteValues

'Q10d
Report.Activate
Report.Range("B358:F365").Copy
Tmplt.Activate
Tmplt.Worksheets("Q10d").Range("D24").PasteSpecial xlPasteValues

'Q10e
Report.Activate
Report.Range("B374:F381").Copy
Tmplt.Activate
Tmplt.Worksheets("Q10e").Range("D24").PasteSpecial xlPasteValues

'Q10f
Report.Activate
Report.Range("B390:F397").Copy
Tmplt.Activate
Tmplt.Worksheets("Q10f").Range("D24").PasteSpecial xlPasteValues
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
    Data.Range("A2:N21").Copy
    Data.Cells(Rp, 17).PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    'Copy & Transpose Paste ends here
    
    'Shifting cells up module starts here
      Data.Range("A1:N24").Delete Shift:=xlUp
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
    Data.Activate
    R2 = Cells(R1, 18).End(xlDown).Row
    
  
    Data.Range(Cells(R1, 18), Cells(R2, 18)).Select
    Selection.Copy
    Report.Activate
    Report.Cells(R1, 1).PasteSpecial xlPasteAll
    
    Data.Activate
    j = 3
    For X = 20 To 29
        N = X - 1 'For calculating the sample size for each Question
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
Data.Range(Cells(R1, 36), Cells(R2, 36)).Select
Selection.Copy
Report.Activate
Report.Cells(R1, 2).PasteSpecial xlPasteAll
Report.Cells(R1 - 1, 2).FormulaR1C1 = "=SUM(RC[1]:RC[4])"

R1 = R1 + 16
Data.Cells(R1, 17).Font.Color = vbGreen

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
Report.Range("B6:F11").Copy
Tmplt.Activate
Tmplt.Worksheets("Q1").Range("C34").PasteSpecial xlPasteValues

'Q2
Report.Range("B22:F25").Copy
Tmplt.Activate
Tmplt.Worksheets("Q2").Range("C30").PasteSpecial xlPasteValues

'Q3 Base
Report.Activate
Report.Range("B38:F38").Copy
Tmplt.Activate
Tmplt.Worksheets("Q3").Range("C38").PasteSpecial xlPasteValues

    'Q3a eBay Money Back Guarantee
    Report.Range("B40:F40").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("C39").PasteSpecial xlPasteValues
    
    'Q3b Trust in the seller
    Report.Range("B56:F56").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("C40").PasteSpecial xlPasteValues

    'Q3c Detailed item description
    Report.Range("B72:F72").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("C41").PasteSpecial xlPasteValues

    'Q3d High quality images that show item details
    Report.Range("B88:F88").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("C42").PasteSpecial xlPasteValues

    'Q3e Returns policy enabled me to return the item if I didn’t like it
    Report.Range("B104:F104").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("C43").PasteSpecial xlPasteValues

    'Q3f Listing stated the item was authentic
    Report.Range("B120:F120").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("C44").PasteSpecial xlPasteValues

    'Q3g Others (Please Specify)
    Report.Range("B136:F136").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("C45").PasteSpecial xlPasteValues

'Q4a Base
Report.Activate
Report.Range("B150:F150").Copy
Tmplt.Activate
Tmplt.Worksheets("Q4a").Range("C34").PasteSpecial xlPasteValues

    'Q4 Extremely Likely -  5
    Report.Range("B155:F155").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4a").Range("C35").PasteSpecial xlPasteValues

    'Q4                     4
    Report.Range("B154:F154").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4a").Range("C36").PasteSpecial xlPasteValues

    'Q4                     3
    Report.Range("B153:F153").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4a").Range("C37").PasteSpecial xlPasteValues

    'Q4                     2
    Report.Range("B152:F152").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4a").Range("C38").PasteSpecial xlPasteValues

    'Q4 Extremely Unlikely - 1
    Report.Range("B151:F151").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4a").Range("C39").PasteSpecial xlPasteValues


'Q4b Base
Report.Activate
Report.Range("B166:F166").Copy
Tmplt.Activate
Tmplt.Worksheets("Q4b").Range("C34").PasteSpecial xlPasteValues

    'Q4 Extremely Likely -  5
    Report.Range("B171:F171").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4b").Range("C35").PasteSpecial xlPasteValues

    'Q4                     4
    Report.Range("B170:F170").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4b").Range("C36").PasteSpecial xlPasteValues

    'Q4                     3
    Report.Range("B169:F169").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4b").Range("C37").PasteSpecial xlPasteValues

    'Q4                     2
    Report.Range("B168:F168").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4b").Range("C38").PasteSpecial xlPasteValues

    'Q4 Extremely Unlikely - 1
    Report.Range("B167:F167").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4b").Range("C39").PasteSpecial xlPasteValues


'Q5
Report.Activate
Report.Range("B182:F191").Copy
Tmplt.Activate
Tmplt.Worksheets("Q5").Range("C42").PasteSpecial xlPasteValues

'Q6
Report.Activate
Report.Range("B198:F198").Copy
Tmplt.Activate
Tmplt.Worksheets("Q6").Range("C34").PasteSpecial xlPasteValues

    'Q6 Extremely Likely -  5
    Report.Range("B203:F203").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q6").Range("C35").PasteSpecial xlPasteValues

    'Q6                     4
    Report.Range("B202:F202").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q6").Range("C36").PasteSpecial xlPasteValues

    'Q6                     3
    Report.Range("B201:F201").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q6").Range("C37").PasteSpecial xlPasteValues

    'Q6                     2
    Report.Range("B200:F200").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q6").Range("C38").PasteSpecial xlPasteValues

    'Q6 Extremely Unlikely - 1
    Report.Range("B199:F199").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q6").Range("C39").PasteSpecial xlPasteValues
    
'Q7
Report.Activate
Report.Range("B214:F214").Copy
Tmplt.Activate
Tmplt.Worksheets("Q7").Range("C34").PasteSpecial xlPasteValues

    'Q7 Extremely Likely -  5
    Report.Range("B219:F219").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q7").Range("C35").PasteSpecial xlPasteValues

    'Q7                     4
    Report.Range("B218:F218").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q7").Range("C36").PasteSpecial xlPasteValues

    'Q7                     3
    Report.Range("B217:F217").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q7").Range("C37").PasteSpecial xlPasteValues

    'Q7                     2
    Report.Range("B216:F216").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q7").Range("C38").PasteSpecial xlPasteValues

    'Q7 Extremely Unlikely - 1
    Report.Range("B215:F215").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q7").Range("C39").PasteSpecial xlPasteValues

'Q8
Report.Activate
Report.Range("B230:F235").Copy
Tmplt.Activate
Tmplt.Worksheets("Q8").Range("C34").PasteSpecial xlPasteValues


'Q9 Base
Report.Activate
Report.Range("B246:F246").Copy
Tmplt.Activate
Tmplt.Worksheets("Q9").Range("C32").PasteSpecial xlPasteValues

    'Q9a eBay brand / product experts
    Report.Range("B248:F248").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q9").Range("C33").PasteSpecial xlPasteValues
    
    'Q9b Professional brand / product experts
    Report.Range("B264:F264").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q9").Range("C34").PasteSpecial xlPasteValues

    'Q9c eBay Professional sellers who have expertise
    Report.Range("B280:F280").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q9").Range("C35").PasteSpecial xlPasteValues

    'Q9d None of the above
    Report.Range("B296:F296").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q9").Range("C36").PasteSpecial xlPasteValues
    
'Q10a
Report.Activate
Report.Range("B310:F317").Copy
Tmplt.Activate
Tmplt.Worksheets("Q10a").Range("D38").PasteSpecial xlPasteValues

'Q10b
Report.Activate
Report.Range("B326:F333").Copy
Tmplt.Activate
Tmplt.Worksheets("Q10b").Range("D38").PasteSpecial xlPasteValues

'Q10c
Report.Activate
Report.Range("B342:F349").Copy
Tmplt.Activate
Tmplt.Worksheets("Q10c").Range("D38").PasteSpecial xlPasteValues

'Q10d
Report.Activate
Report.Range("B358:F365").Copy
Tmplt.Activate
Tmplt.Worksheets("Q10d").Range("D38").PasteSpecial xlPasteValues

'Q10e
Report.Activate
Report.Range("B374:F381").Copy
Tmplt.Activate
Tmplt.Worksheets("Q10e").Range("D38").PasteSpecial xlPasteValues

'Q10f
Report.Activate
Report.Range("B390:F397").Copy
Tmplt.Activate
Tmplt.Worksheets("Q10f").Range("D38").PasteSpecial xlPasteValues
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
    Data.Range("A2:N21").Copy
    Data.Cells(Rp, 17).PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    'Copy & Transpose Paste ends here
    
    'Shifting cells up module starts here
      Data.Range("A1:N24").Delete Shift:=xlUp
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
For Q = 1 To 25 'For UK & US change 18 to 25
    Data.Activate
    R2 = Cells(R1, 18).End(xlDown).Row
    
  
    Data.Range(Cells(R1, 18), Cells(R2, 18)).Select
    Selection.Copy
    Report.Activate
    Report.Cells(R1, 1).PasteSpecial xlPasteAll
    
    Data.Activate
    j = 3
    For X = 20 To 29
        N = X - 1 'For calculating the sample size for each Question
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
Data.Range(Cells(R1, 36), Cells(R2, 36)).Select
Selection.Copy
Report.Activate
Report.Cells(R1, 2).PasteSpecial xlPasteAll
Report.Cells(R1 - 1, 2).FormulaR1C1 = "=SUM(RC[1]:RC[4])"

R1 = R1 + 16
Data.Cells(R1, 17).Font.Color = vbGreen

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
Report.Range("B6:F11").Copy
Tmplt.Activate
Tmplt.Worksheets("Q1").Range("C46").PasteSpecial xlPasteValues

'Q2
Report.Range("B22:F25").Copy
Tmplt.Activate
Tmplt.Worksheets("Q2").Range("C40").PasteSpecial xlPasteValues

'Q3 Base
Report.Activate
Report.Range("B38:F38").Copy
Tmplt.Activate
Tmplt.Worksheets("Q3").Range("C52").PasteSpecial xlPasteValues

    'Q3a eBay Money Back Guarantee
    Report.Range("B40:F40").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("C53").PasteSpecial xlPasteValues
    
    'Q3b Trust in the seller
    Report.Range("B56:F56").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("C54").PasteSpecial xlPasteValues

    'Q3c Detailed item description
    Report.Range("B72:F72").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("C55").PasteSpecial xlPasteValues

    'Q3d High quality images that show item details
    Report.Range("B88:F88").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("C56").PasteSpecial xlPasteValues

    'Q3e Returns policy enabled me to return the item if I didn’t like it
    Report.Range("B104:F104").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("C57").PasteSpecial xlPasteValues

    'Q3f Listing stated the item was authentic
    Report.Range("B120:F120").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("C58").PasteSpecial xlPasteValues

    'Q3g Others (Please Specify)
    Report.Range("B136:F136").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("C59").PasteSpecial xlPasteValues

'Q4a Base
Report.Activate
Report.Range("B150:F150").Copy
Tmplt.Activate
Tmplt.Worksheets("Q4a").Range("C46").PasteSpecial xlPasteValues

    'Q4 Extremely Likely -  5
    Report.Range("B155:F155").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4a").Range("C47").PasteSpecial xlPasteValues

    'Q4                     4
    Report.Range("B154:F154").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4a").Range("C48").PasteSpecial xlPasteValues

    'Q4                     3
    Report.Range("B153:F153").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4a").Range("C49").PasteSpecial xlPasteValues

    'Q4                     2
    Report.Range("B152:F152").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4a").Range("C50").PasteSpecial xlPasteValues

    'Q4 Extremely Unlikely - 1
    Report.Range("B151:F151").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4a").Range("C51").PasteSpecial xlPasteValues


'Q4b Base
Report.Activate
Report.Range("B166:F166").Copy
Tmplt.Activate
Tmplt.Worksheets("Q4b").Range("C46").PasteSpecial xlPasteValues

    'Q4 Extremely Likely -  5
    Report.Range("B171:F171").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4b").Range("C47").PasteSpecial xlPasteValues

    'Q4                     4
    Report.Range("B170:F170").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4b").Range("C48").PasteSpecial xlPasteValues

    'Q4                     3
    Report.Range("B169:F169").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4b").Range("C49").PasteSpecial xlPasteValues

    'Q4                     2
    Report.Range("B168:F168").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4b").Range("C50").PasteSpecial xlPasteValues

    'Q4 Extremely Unlikely - 1
    Report.Range("B167:F167").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4b").Range("C51").PasteSpecial xlPasteValues


'Q5
Report.Activate
Report.Range("B182:F191").Copy
Tmplt.Activate
Tmplt.Worksheets("Q5").Range("C58").PasteSpecial xlPasteValues

'Q6
Report.Activate
Report.Range("B198:F198").Copy
Tmplt.Activate
Tmplt.Worksheets("Q6").Range("C46").PasteSpecial xlPasteValues

    'Q6 Extremely Likely -  5
    Report.Range("B203:F203").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q6").Range("C47").PasteSpecial xlPasteValues

    'Q6                     4
    Report.Range("B202:F202").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q6").Range("C48").PasteSpecial xlPasteValues

    'Q6                     3
    Report.Range("B201:F201").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q6").Range("C49").PasteSpecial xlPasteValues

    'Q6                     2
    Report.Range("B200:F200").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q6").Range("C50").PasteSpecial xlPasteValues

    'Q6 Extremely Unlikely - 1
    Report.Range("B199:F199").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q6").Range("C51").PasteSpecial xlPasteValues
    
'Q7
Report.Activate
Report.Range("B214:F214").Copy
Tmplt.Activate
Tmplt.Worksheets("Q7").Range("C46").PasteSpecial xlPasteValues

    'Q7 Extremely Likely -  5
    Report.Range("B219:F219").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q7").Range("C47").PasteSpecial xlPasteValues

    'Q7                     4
    Report.Range("B218:F218").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q7").Range("C48").PasteSpecial xlPasteValues

    'Q7                     3
    Report.Range("B217:F217").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q7").Range("C49").PasteSpecial xlPasteValues

    'Q7                     2
    Report.Range("B216:F216").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q7").Range("C50").PasteSpecial xlPasteValues

    'Q7 Extremely Unlikely - 1
    Report.Range("B215:F215").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q7").Range("C51").PasteSpecial xlPasteValues

'Q8
Report.Activate
Report.Range("B230:F235").Copy
Tmplt.Activate
Tmplt.Worksheets("Q8").Range("C46").PasteSpecial xlPasteValues


'Q9 Base
Report.Activate
Report.Range("B246:F246").Copy
Tmplt.Activate
Tmplt.Worksheets("Q9").Range("C43").PasteSpecial xlPasteValues

    'Q9a eBay brand / product experts
    Report.Range("B248:F248").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q9").Range("C44").PasteSpecial xlPasteValues
    
    'Q9b Professional brand / product experts
    Report.Range("B264:F264").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q9").Range("C45").PasteSpecial xlPasteValues

    'Q9c eBay Professional sellers who have expertise
    Report.Range("B280:F280").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q9").Range("C46").PasteSpecial xlPasteValues

    'Q9d None of the above
    Report.Range("B296:F296").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q9").Range("C47").PasteSpecial xlPasteValues
    
'Q10a
Report.Activate
Report.Range("B310:F317").Copy
Tmplt.Activate
Tmplt.Worksheets("Q10a").Range("D52").PasteSpecial xlPasteValues

'Q10b
Report.Activate
Report.Range("B326:F333").Copy
Tmplt.Activate
Tmplt.Worksheets("Q10b").Range("D52").PasteSpecial xlPasteValues

'Q10c
Report.Activate
Report.Range("B342:F349").Copy
Tmplt.Activate
Tmplt.Worksheets("Q10c").Range("D52").PasteSpecial xlPasteValues

'Q10d
Report.Activate
Report.Range("B358:F365").Copy
Tmplt.Activate
Tmplt.Worksheets("Q10d").Range("D52").PasteSpecial xlPasteValues

'Q10e
Report.Activate
Report.Range("B374:F381").Copy
Tmplt.Activate
Tmplt.Worksheets("Q10e").Range("D52").PasteSpecial xlPasteValues

'Q10f
Report.Activate
Report.Range("B390:F397").Copy
Tmplt.Activate
Tmplt.Worksheets("Q10f").Range("D52").PasteSpecial xlPasteValues
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
    Data.Range("A2:N21").Copy
    Data.Cells(Rp, 17).PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    'Copy & Transpose Paste ends here
    
    'Shifting cells up module starts here
      Data.Range("A1:N24").Delete Shift:=xlUp
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
    Data.Activate
    R2 = Cells(R1, 18).End(xlDown).Row
    
  
    Data.Range(Cells(R1, 18), Cells(R2, 18)).Select
    Selection.Copy
    Report.Activate
    Report.Cells(R1, 1).PasteSpecial xlPasteAll
    
    Data.Activate
    j = 3
    For X = 20 To 29
        N = X - 1 'For calculating the sample size for each Question
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
Data.Range(Cells(R1, 36), Cells(R2, 36)).Select
Selection.Copy
Report.Activate
Report.Cells(R1, 2).PasteSpecial xlPasteAll
Report.Cells(R1 - 1, 2).FormulaR1C1 = "=SUM(RC[1]:RC[4])"

R1 = R1 + 16
Data.Cells(R1, 17).Font.Color = vbGreen

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
Report.Range("B6:F11").Copy
Tmplt.Activate
Tmplt.Worksheets("Q1").Range("C10").PasteSpecial xlPasteValues

'Q2
Report.Range("B22:F25").Copy
Tmplt.Activate
Tmplt.Worksheets("Q2").Range("C10").PasteSpecial xlPasteValues

'Q3 Base
Report.Activate
Report.Range("B38:F38").Copy
Tmplt.Activate
Tmplt.Worksheets("Q3").Range("C10").PasteSpecial xlPasteValues

    'Q3a eBay Money Back Guarantee
    Report.Range("B40:F40").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("C11").PasteSpecial xlPasteValues
    
    'Q3b Trust in the seller
    Report.Range("B56:F56").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("C12").PasteSpecial xlPasteValues

    'Q3c Detailed item description
    Report.Range("B72:F72").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("C13").PasteSpecial xlPasteValues

    'Q3d High quality images that show item details
    Report.Range("B88:F88").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("C14").PasteSpecial xlPasteValues

    'Q3e Returns policy enabled me to return the item if I didn’t like it
    Report.Range("B104:F104").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("C15").PasteSpecial xlPasteValues

    'Q3f Listing stated the item was authentic
    Report.Range("B120:F120").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("C16").PasteSpecial xlPasteValues

    'Q3g Others (Please Specify)
    Report.Range("B136:F136").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q3").Range("C17").PasteSpecial xlPasteValues

'Q4a Base
Report.Activate
Report.Range("B150:F150").Copy
Tmplt.Activate
Tmplt.Worksheets("Q4a").Range("C10").PasteSpecial xlPasteValues

    'Q4 Extremely Likely -  5
    Report.Range("B155:F155").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4a").Range("C11").PasteSpecial xlPasteValues

    'Q4                     4
    Report.Range("B154:F154").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4a").Range("C12").PasteSpecial xlPasteValues

    'Q4                     3
    Report.Range("B153:F153").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4a").Range("C13").PasteSpecial xlPasteValues

    'Q4                     2
    Report.Range("B152:F152").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4a").Range("C14").PasteSpecial xlPasteValues

    'Q4 Extremely Unlikely - 1
    Report.Range("B151:F151").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4a").Range("C15").PasteSpecial xlPasteValues


'Q4b Base
Report.Activate
Report.Range("B166:F166").Copy
Tmplt.Activate
Tmplt.Worksheets("Q4b").Range("C10").PasteSpecial xlPasteValues

    'Q4 Extremely Likely -  5
    Report.Range("B171:F171").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4b").Range("C11").PasteSpecial xlPasteValues

    'Q4                     4
    Report.Range("B170:F170").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4b").Range("C12").PasteSpecial xlPasteValues

    'Q4                     3
    Report.Range("B169:F169").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4b").Range("C13").PasteSpecial xlPasteValues

    'Q4                     2
    Report.Range("B168:F168").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4b").Range("C14").PasteSpecial xlPasteValues

    'Q4 Extremely Unlikely - 1
    Report.Range("B167:F167").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q4b").Range("C15").PasteSpecial xlPasteValues


'Q5
Report.Activate
Report.Range("B182:F191").Copy
Tmplt.Activate
Tmplt.Worksheets("Q5").Range("C10").PasteSpecial xlPasteValues

'Q6
Report.Activate
Report.Range("B198:F198").Copy
Tmplt.Activate
Tmplt.Worksheets("Q6").Range("C10").PasteSpecial xlPasteValues

    'Q6 Extremely Likely -  5
    Report.Range("B203:F203").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q6").Range("C11").PasteSpecial xlPasteValues

    'Q6                     4
    Report.Range("B202:F202").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q6").Range("C12").PasteSpecial xlPasteValues

    'Q6                     3
    Report.Range("B201:F201").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q6").Range("C13").PasteSpecial xlPasteValues

    'Q6                     2
    Report.Range("B200:F200").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q6").Range("C14").PasteSpecial xlPasteValues

    'Q6 Extremely Unlikely - 1
    Report.Range("B199:F199").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q6").Range("C15").PasteSpecial xlPasteValues
    
'Q7
Report.Activate
Report.Range("B214:F214").Copy
Tmplt.Activate
Tmplt.Worksheets("Q7").Range("C10").PasteSpecial xlPasteValues

    'Q7 Extremely Likely -  5
    Report.Range("B219:F219").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q7").Range("C11").PasteSpecial xlPasteValues

    'Q7                     4
    Report.Range("B218:F218").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q7").Range("C12").PasteSpecial xlPasteValues

    'Q7                     3
    Report.Range("B217:F217").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q7").Range("C13").PasteSpecial xlPasteValues

    'Q7                     2
    Report.Range("B216:F216").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q7").Range("C14").PasteSpecial xlPasteValues

    'Q7 Extremely Unlikely - 1
    Report.Range("B215:F215").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q7").Range("C15").PasteSpecial xlPasteValues

'Q8
Report.Activate
Report.Range("B230:F235").Copy
Tmplt.Activate
Tmplt.Worksheets("Q8").Range("C10").PasteSpecial xlPasteValues


'Q9 Base
Report.Activate
Report.Range("B246:F246").Copy
Tmplt.Activate
Tmplt.Worksheets("Q9").Range("C10").PasteSpecial xlPasteValues

    'Q9a eBay brand / product experts
    Report.Range("B248:F248").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q9").Range("C11").PasteSpecial xlPasteValues
    
    'Q9b Professional brand / product experts
    Report.Range("B264:F264").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q9").Range("C12").PasteSpecial xlPasteValues

    'Q9c eBay Professional sellers who have expertise
    Report.Range("B280:F280").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q9").Range("C13").PasteSpecial xlPasteValues

    'Q9d None of the above
    Report.Range("B296:F296").Copy
    Tmplt.Activate
    Tmplt.Worksheets("Q9").Range("C14").PasteSpecial xlPasteValues
    
'Q10a
Report.Activate
Report.Range("B310:F317").Copy
Tmplt.Activate
Tmplt.Worksheets("Q10a").Range("D10").PasteSpecial xlPasteValues

'Q10b
Report.Activate
Report.Range("B326:F333").Copy
Tmplt.Activate
Tmplt.Worksheets("Q10b").Range("D10").PasteSpecial xlPasteValues

'Q10c
Report.Activate
Report.Range("B342:F349").Copy
Tmplt.Activate
Tmplt.Worksheets("Q10c").Range("D10").PasteSpecial xlPasteValues

'Q10d
Report.Activate
Report.Range("B358:F365").Copy
Tmplt.Activate
Tmplt.Worksheets("Q10d").Range("D10").PasteSpecial xlPasteValues

'Q10e
Report.Activate
Report.Range("B374:F381").Copy
Tmplt.Activate
Tmplt.Worksheets("Q10e").Range("D10").PasteSpecial xlPasteValues

'Q10f
Report.Activate
Report.Range("B390:F397").Copy
Tmplt.Activate
Tmplt.Worksheets("Q10f").Range("D10").PasteSpecial xlPasteValues
End Sub

