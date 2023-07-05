Option Explicit
Global wb As Workbook
Global sh As Worksheet
Global tkr_rng As Double
Global tkr As String, gtst_inc_tkr As String, gtst_tsv_tkr As String, gtst_dec_tkr As String
Global i As Double, opn As Double, cls As Double, yr_chg As Double, prct_chg As Double
Global tsv As Double, col_L As Long, LastRow As Long, gtst_inc As Double, gtst_dec As Double, gtst_tsv As Double
Global Ticker As String, Yearly_Change As String, Percentage_Change As String, Total_Stock_Volume As String
Global Greatest_Increase As String, Greatest_Decrease As String, Greatest_Volume As String


Sub TkrNm()

For Each sh In ActiveWorkbook.Worksheets
'sh.Activate

'initialise variables
LastRow = sh.Cells(Rows.Count, 1).End(xlUp).Row

opn = sh.Cells(2, 3).Value

tkr_rng = 1
tsv = 0
yr_chg = 0
gtst_tsv = 0
gtst_inc = 0
gtst_dec = 0
gtst_tsv_tkr = " "
gtst_inc_tkr = " "
gtst_dec_tkr = " "


For i = 2 To LastRow

If sh.Cells(i + 1, 1).Value <> sh.Cells(i, 1).Value Then
tkr = sh.Cells(i, 1).Value
tsv = tsv + sh.Cells(i + 1, 7).Value
Else
tsv = tsv + sh.Cells(i + 1, 7).Value
                        
    If tsv > gtst_tsv Then
    gtst_tsv = tsv
    gtst_tsv_tkr = tkr
    End If
   
    cls = sh.Cells(i, 6).Value
    yr_chg = cls - opn

                
    If opn <> 0 Then
    prct_chg = yr_chg / opn
    End If
                
    tkr_rng = tkr_rng + 1
                
    If prct_chg > gtst_inc Then
    gtst_inc = prct_chg
    gtst_inc_tkr = tkr
    End If
                
    If prct_chg < gtst_dec Then
    gtst_dec = prct_chg
    gtst_dec_tkr = tkr
    End If
                         

sh.Range("I" & tkr_rng).Value = tkr
sh.Range("L" & tkr_rng).Value = tsv
sh.Range("J" & tkr_rng).Value = yr_chg
sh.Range("O2").Value = gtst_inc_tkr
sh.Range("O3").Value = gtst_dec_tkr = tkr
sh.Range("O4").Value = gtst_tsv_tkr
sh.Range("P2").Value = gtst_inc
sh.Range("P3").Value = gtst_dec
sh.Range("P4").Value = gtst_tsv
    
    If sh.Cells(i, 10) >= 0 Then
    sh.Cells(i, 10).Interior.ColorIndex = 4
    Else
    sh.Cells(i, 10).Interior.ColorIndex = 3
    End If
    
'Set Headings
sh.Range("I1").Value = "Ticker"
sh.Range("J1").Value = "Yearly_Change"
sh.Range("K1").Value = "Percentage_Change"
sh.Range("L1").Value = "Total_Stock_Volume"
sh.Range("O1").Value = "Ticker"
sh.Range("P1").Value = "Value"
sh.Range("N2").Value = "Greatest_Increase"
sh.Range("N3").Value = "Greatest_Decrease"
sh.Range("N4").Value = "Greatest_Volume"
sh.Range("K" & tkr_rng).Value = prct_chg
sh.Range("K" & tkr_rng).NumberFormat = "0.00%"
sh.Range("L:L").ColumnWidth = 15
sh.Range("P:P").ColumnWidth = 15

End If

Next i

Next sh

End Sub




