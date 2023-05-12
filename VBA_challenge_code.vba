{\rtf1\ansi\ansicpg1252\cocoartf2709
\cocoatextscaling0\cocoaplatform0{\fonttbl\f0\fswiss\fcharset0 Helvetica;}
{\colortbl;\red255\green255\blue255;}
{\*\expandedcolortbl;;}
\margl1440\margr1440\vieww11520\viewh8400\viewkind0
\pard\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\pardirnatural\partightenfactor0

\f0\fs24 \cf0 Sub Stockprice_1()\
\
     ' Declare all necessary variables\
  Dim Ticker As String\
  Dim Opening_price As Double\
  Dim Closing_price As Double\
  Dim Yearly_change As Double\
  Dim Percentage_change As Double\
  Dim Total_volume As Double\
  Dim Summary_Table_Row As Long ' Use Long instead of Integer to avoid overflow errors\
  Dim Lastrow As Double 'go to the last row of the list\
  Dim ws As Worksheet 'to itterate between the worksheets\
  \
For Each ws In Worksheets\
\
  ' Set initial values for variables\
  Total_volume = 0\
  Summary_Table_Row = 2 ' Start on row 2 to leave space for header\
  Start = 2\
 \
  ' Add headers to summary table\
  ws.Range("I1").Value = "Ticker"\
  ws.Range("J1").Value = "Yearly Change"\
  ws.Range("K1").Value = "Percentage Change"\
  ws.Range("L1").Value = "Total Volume"\
  ws.Range("O2") = "Greatest % Increase Value"\
  ws.Range("O3") = "Greatest % Decrease Value"\
  ws.Range("O4") = "Greatest Total Volume"\
  ws.Range("P1") = "Ticker"\
  ws.Range("Q1") = "Value"\
  \
  ' Add "Bold" format to header\
  ws.Range("I1").Font.Bold = True\
  ws.Range("J1").Font.Bold = True\
  ws.Range("K1").Font.Bold = True\
  ws.Range("L1").Font.Bold = True\
  ws.Range("O2").Font.Bold = True\
  ws.Range("O3").Font.Bold = True\
  ws.Range("O4").Font.Bold = True\
  ws.Range("P1").Font.Bold = True\
  ws.Range("Q1").Font.Bold = True\
\
\
Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row\
\
  ' Loop through all stock prices\
  For i = 2 To Lastrow ' data from 2nd row till last row\
  \
      'If previous ticker and current ticker are not the same, then\
        If ws.Cells(i - 1, 1) <> ws.Cells(i, 1) Then\
        Opening_price = ws.Cells(i, 3)\
\
    ' Check if it is at the same stock ticker, if not\
    ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then\
\
      ' Set the Ticker name\
      Ticker = ws.Cells(i, 1).Value\
\
      ' Add to the total volume\
        Total_volume = Total_volume + ws.Cells(i, 7).Value\
      \
        'Set the closing price\
        Closing_price = ws.Cells(i, 6).Value\
      \
      ' Print the ticker and total volume in the summary table\
      ws.Range("I" & Summary_Table_Row).Value = Ticker\
      ws.Range("L" & Summary_Table_Row).Value = Total_volume\
\
      ' Calculate the yearly change and percentage change for the ticker\
      Yearly_change = Closing_price - Opening_price\
      Percentage_change = (Yearly_change / Opening_price) * 100\
      On Error Resume Next\
      \
      ' Print the yearly change and percentage change in the summary table\
      ws.Range("J" & Summary_Table_Row).Value = Yearly_change\
      ws.Range("K" & Summary_Table_Row).Value = Percentage_change\
      ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00\\%"\
      \
      ' Add the summary table row and reset the total volume\
      Summary_Table_Row = Summary_Table_Row + 1\
      Total_volume = 0\
\
    ' If the cell immediately following a row is the same ticker...\
    Else\
\
      ' Add to the total volume for the ticker\
      Total_volume = Total_volume + ws.Cells(i, 7).Value\
\
    End If\
\
  Next i\
\
        \
    \
   'After the 1st loop is done, set the next loop\
       \
       greatest_increase = ws.Cells(2, 11)\
       greatest_decrease = ws.Cells(2, 11)\
       greatest_volume = ws.Cells(2, 12)\
       Lastrow_summary = ws.Cells(Rows.Count, 10).End(xlUp).Row\
  \
       For j = 2 To Lastrow_summary\
       \
        'Change the format depending on the value\
        If ws.Cells(j, 10) >= 0 Then\
        ws.Cells(j, 10).Interior.ColorIndex = 4\
  \
        ElseIf ws.Cells(j, 10) < 0 Then\
        ws.Cells(j, 10).Interior.ColorIndex = 3\
   \
        End If\
        \
        'Loop through each row and replace the greatest increase value\
        If ws.Cells(j, 11) > greatest_increase Then\
        greatest_increase = ws.Cells(j, 11)\
        ws.Cells(2, 17) = greatest_increase\
        ws.Cells(2, 17).NumberFormat = "0.00\\%" 'for percentage format\
        ws.Cells(2, 16) = ws.Cells(j, 9)\
   \
        End If\
   \
        'Loop through each row and replace the greatest decrease value\
        If ws.Cells(j, 11) < greatest_decrease Then\
        greatest_decrease = ws.Cells(j, 11)\
        ws.Cells(3, 17) = greatest_decrease\
        ws.Cells(3, 17).NumberFormat = "0.00\\%" 'for percentage format\
        ws.Cells(3, 16) = ws.Cells(j, 9)\
  \
        End If\
        \
        'Loop through each row and replace the greatest total volume\
        If ws.Cells(j, 12) > greatest_volume Then\
        greatest_volume = ws.Cells(j, 12)\
        ws.Cells(4, 17) = greatest_volume\
        ws.Cells(4, 16) = ws.Cells(j, 9)\
   \
        End If\
   \
       Next j\
       \
    ws.Columns("I:Q").AutoFit ' Autofit the cell as per the content in the cell\
  \
Next ws\
\
\
End Sub\
\
}