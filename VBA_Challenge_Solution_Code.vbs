{\rtf1\ansi\ansicpg1252\cocoartf2709
\cocoatextscaling0\cocoaplatform0{\fonttbl\f0\fswiss\fcharset0 Helvetica;}
{\colortbl;\red255\green255\blue255;}
{\*\expandedcolortbl;;}
\paperw11900\paperh16840\margl1440\margr1440\vieww11520\viewh8400\viewkind0
\pard\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\pardirnatural\partightenfactor0

\f0\fs24 \cf0 Option Explicit\
Global wb As Workbook\
Global sh As Worksheet\
Global tkr_rng As Integer, max As Double, min As Double\
Global tkr As String, st_tkr As String, nx_tkr As String, Ticker As String\
Global Yearly_Change As String, Percentage_Change As String, Total_Stock_Volume As String\
Global Greatest_Increase As String, Greatest_Decrease As String, Greatest_Volume As String\
Global opn As Double, cls As Double, yr_chg As Double, prct_chg As Double\
Global ttl_vol As Double, tsv As Double, col_L As Double, opn_vol As Double\
Global i As Long, LastRow As Long\
\
\
Global tkr_rng1 As Range\
\
Sub TkrNm()\
\
For Each sh In ActiveWorkbook.Worksheets\
\
'\'91Set Headings\
sh.Range("I1").Value = "Ticker"\
sh.Range("J1").Value = "Yearly_Change"\
sh.Range("K1").Value = "Percentage_Change"\
sh.Range("L1").Value = "Total_Stock_Volume"\
sh.Range("O1").Value = "Ticker"\
sh.Range("P1").Value = "Value"\
sh.Range("N2").Value = "Greatest_Increase"\
sh.Range("N3").Value = "Greatest_Decrease"\
sh.Range("N4").Value = "Greatest_Volume"\
\
'Populate summary table\
tkr_rng = 2\
tsv = 0\
yr_chg = 0\
opn = 0\
cls = 0\
\
LastRow = sh.Cells(Rows.Count, 1).End(xlUp).Row\
\
For i = 2 To LastRow\
\
            If sh.Cells(i + 1, 1).Value <> sh.Cells(i, 1).Value Then\
                tkr = sh.Cells(i, 1).Value\
                tsv = tsv + sh.Cells(i + 1, 7).Value\
                opn = sh.Cells(i + 1, 3).Value\
                cls = sh.Cells(i + 1, 6).Value\
                yr_chg = cls - opn\
                opn_vol = sh.Cells(i + 1, 7).Value\
         \
        sh.Range("I" & tkr_rng).Value = tkr\
        sh.Range("L" & tkr_rng).Value = tsv\
       sh.Range("J" & tkr_rng).Value = yr_chg\
\
        tkr_rng = tkr_rng + 1\
\
      tsv = 0\
      yr_chg = 0\
    \
   Else\
      tsv = tsv + sh.Cells(i, 7).Value\
       \
    End If\
    \
    If opn <> 0 Then\
    \
    prct_chg = (cls - opn) / opn\
    \
    \
Else\
 prct_chg = 0\
\
End If\
    \
        sh.Range("K" & tkr_rng).Value = prct_chg\
        sh.Range("K" & tkr_rng).NumberFormat = "0.00%"\
        sh.Range("L:L").ColumnWidth = 15\
        sh.Range("N:N").ColumnWidth = 15\
        sh.Range("P:P").ColumnWidth = 15\
\
   End If\
\
\
'\'91Calculate greatest increase\
\
If Cells(i + 1, 11).Value > Cells(i, 11).Value Then\
    max = sh.Cells(i, 11).Value\
    tkr = sh.Cells(i, 9).Value\
    sh.Range("P2").Value = max\
    sh.Range("O2").Value = tkr\
    sh.Range("P2").NumberFormat = "0.00%"\
\
End If\
\
If sh.Cells(i + 1, 12).Value < sh.Cells(i, 12).Value Then\
    col_L = sh.Cells(i, 12).Value\
    tkr = sh.Cells(i, 9).Value\
    sh.Range("P4").Value = col_L\
    sh.Range("O4").Value = tkr\
\
\
End If\
\
If sh.Cells(i + 1, 12).Value > sh.Cells(i, 12).Value Then\
    min = sh.Cells(i, 11).Value\
    tkr = sh.Cells(i, 9).Value\
    sh.Range("P3").Value = min\
    sh.Range("O3").Value = tkr\
    sh.Range("P3").NumberFormat = "0.00%"\
\
End If\
\
If Cells(i, 10) >= 0 Then\
\
        sh.Cells(i, 10).Interior.ColorIndex = 4\
Else\
        sh.Cells(i, 10).Interior.ColorIndex = 3\
End If\
\
Next i\
\
Next sh\
\
End Sub\
\
\
\
}