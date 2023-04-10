{\rtf1\ansi\ansicpg1252\cocoartf2709
\cocoatextscaling0\cocoaplatform0{\fonttbl\f0\fswiss\fcharset0 Helvetica;}
{\colortbl;\red255\green255\blue255;}
{\*\expandedcolortbl;;}
\margl1440\margr1440\vieww15760\viewh14220\viewkind0
\pard\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\pardirnatural\partightenfactor0

\f0\fs24 \cf0 \'91. Aaron Otto 4.9.23 UofM Data Boot Camp Challenge 2 Code file \
\'91 ****REQUIREMENTS*******\
\'91Looping Across Worksheet (20 points)\
\
\'91This first sub \'93SubMultiSheet\'94 can be called to process all three pages in the workbook.  \
\'91 It will call the first process then change the active sheet \
\'91 Cycling through all three sheets \
\'91 Please either copy and paste the code below the line into a module in a workbook you would like to run\
\'91 Or change extension of this file to .VBS and save \
\
\
Sub CallMultiSheets()\
\
    ' Call the sub procedure on the first worksheet while its active \
    Worksheets("2018").Activate  \
    Call StockCounter  \
    \
    ' Call the sub procedure on the second worksheet\
    Worksheets("2019").Activate  \
    Call StockCounter  \
    \
    ' Call the sub procedure on the third worksheet\
    Worksheets("2020").Activate  \
    Call StockCounter  \
End Sub\
\
\'91 **************************************************************************************************\
\'91\
\'91\
\'91\
\
Sub StockCounter()\
\
' This program will perform the following activities.  Each activity is broken up into main sections according to requirements.\
' Section 1 Delcaing Variables and values\
' Section 2 Adminstrative column formating and header writing. With the exception of the requirement to change interior colors of cell.\
' Section 3 is filling the TickerName Column(I) and populating the YearlyDiff Column(J)\
'    is populating the PercentChange & TotalStockVolumes columns (columns K,L rspectively)\
' Section 4 is Finding the Greatest % increase and poulating cells P2,Q2 with tickername and value\
' Section 5 is Finding the Greatest % decrease and poulating cells P3,Q3 with tickername and value\
' Section 6 is Finding the Greatest Total Volume  cells P4,Q4 with tickername and value\
\
'Section 1 -Declaring variables for process\
\
\
  Dim stockName  As String\
  Dim i As Long\
  Dim TotVol As Double\
  Dim OpenPrice As Double\
  Dim ClosePrice As Double\
  Dim PriceDiff As Double\
  Dim Summary_Table_Row As Integer\
  Dim PerChange As Double\
  Dim lastRow As Long\
  Dim lastRow2 As Long\
\
'Section 2 - Adminstrative\
    'Adding headers for summary table\
    Range("I1").Value = "TickerName"\
    Range("J1").Value = "YearlyChange"\
    Range("K1").Value = "PercentChange"\
    Range("L1").Value = "TotalStockVolume"\
    \
    'Adding column and row headers for Greatest Inc,Dec, and volumes\
    Range("P1").Value = "Ticker"\
    Range("Q1").Value = "Value"\
    Range("O2").Value = "Greatest % Increase"\
    Range("O3").Value = "Greates % Decrease"\
    Range("O4").Value = "Geatest Total Volumes"\
    lastRow = Range("A" & Rows.Count).End(xlUp).Row ' building last row for our loop so we can use this on each spreadsheet\
    Range("J2:J" & lastRow).NumberFormat = "0.00" ' Number formatting.\
    Range("K2:K" & lastRow).NumberFormat = "0.00%" ' % formatting\
    Range("Q2:Q3").NumberFormat = "0.00%" ' % formatting\
\
\'91 ****REQUIREMENTS*******\
\'91Conditional formatting is applied correctly and appropriately to the yearly change column (10 points)\
\'91Conditional formatting is applied correctly and appropriately to the percent change column (10 points)\
\'91Calculated Values (15 points)\
\
    \
 'Section 3 - Looping and summarizing differences, percent change and total stock volumes\
\
\'91 ****REQUIREMENTS*******\
\'91The script loops through one year of stock data and reads/ stores all of the following values from each row:\
\'91ticker symbol (5 points)\
\'91volume of stock (5 points)\
\'91open price (5 points)\
\'91close price (5 points)\
\'91Column Creation (10 points)\
\'91On the same worksheet as the raw data, or on a new worksheet all columns were correctly created for:\
\'91ticker symbol (2.5 points)\
\'91total stock volume (2.5 points)\
\'91yearly change ($) (2.5 points)\
\'91percent change (2.5 points)\
\'91Conditional Formatting (20 points)\
 \
        OpenPrice = Cells(2, 3).Value ' Prior to entering loop setting first OpenPrice value\
        Summary_Table_Row = 2 ' Setting the row position for our summary table.\
       ' lastRow = Range("A" & Rows.Count).End(xlUp).Row ' building last row for our loop so we can use this on each spreadsheet\
        \
        lastRow2 = Cells(Rows.Count, "K").End(xlUp).Row\
\
         For i = 2 To lastRow 'starting loop\
\
                 If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then 'setting logic conditions for loop\
\
                    stockName = Cells(i, 1).Value 'allowing the loop to catch StockName\
                    ClosePrice = Cells(i, 6).Value 'this allows the Close price to be caught by the loop each cycle.\
\
                    PriceDiff = ClosePrice - OpenPrice 'Setting PriceDiff variable\
                    PerChange = PriceDiff / OpenPrice 'Setting Percent Change variable\
\
                ' Print values to summary tables mapped target cells\
                    Range("I" & Summary_Table_Row).Value = stockName\
                    Range("J" & Summary_Table_Row).Value = PriceDiff '\
                    Range("K" & Summary_Table_Row).Value = PerChange\
                    Range("L" & Summary_Table_Row).Value = TotVol\
                    \
                ' Reseting and incrementing variables & ensuring summary table doesn't have values overwritten\
                OpenPrice = Cells(i + 1, 3).Value\
                Summary_Table_Row = Summary_Table_Row + 1\
                TotVol = 0\
\
         Else\
\
            TotVol = TotVol + Cells(i, 7).Value ' If not the right volume we just keep adding.\
\
        End If\
\
  Next i\
\
\
' Section 3 Color Indexes updated for YearlyDifference column which is Column J\
For i = 2 To lastRow\
\
    If Cells(i, 10) > 0 Then ' logic to check each value in the column for being larger then 0\
\
        Cells(i, 10).Interior.ColorIndex = 4  ' if larger then 0 change interior color to green\
\
        Else\
\
        Cells(i, 10).Interior.ColorIndex = 3  'if not larger then zero we change interior color to red\
\
        End If\
\
    Next i\
\
'Section 4 - Finding and populating the Greatest %  Increase\
    Dim largestValue As Double\
    Dim largestValueStock As Variant\
    Dim currentValue As Double\
    Dim currentStock As Variant\
    Dim smallestValue As Double\
\
\'91All three of the following values are calculated correctly and displayed in the output:\
\
\'91\'91 ****REQUIREMENTS******* \
\'91Greatest % Increase (5 points)\
    \
 ' going after largest value in PercentChange column K\
       largestValue = 0\
    \
    \
        For i = 2 To 3001 ' Assuming your data starts in row 2 and you want to search 3000 rows(throwing mismatch error with end formula\
        currentValue = Range("K" & i).Value ' Assuming the column you want to search is column K\
        currentStock = Range("I" & i).Value ' same  assumption here\
        If currentValue > largestValue Then\
            largestValue = currentValue\
           Range("Q2").Value = largestValue\
           Range("P2").Value = currentStock\
\
        End If\
    Next i\
\
   \
    \
    'Section 5 -going after smallest value in Percent change column K ie biggest decrease\
    \
\'91\'91 ****REQUIREMENTS*******\
\'91Greatest % Decrease (5 points)\
 \'91\
       smallestValue = 0\
      \
    For i = 2 To 3001  ' data starts at row 2 but need to define names for all\
             currentValue = Range("K" & i).Value ' seaching PercentChange column (k)\
              currentStock = Range("I" & i).Value ' same  assumption here for column I\
            If currentValue < smallestValue Then\
              '\
              smallestValue = currentValue\
              Range("Q3").Value = smallestValue\
           Range("P3").Value = currentStock\
\
        End If\
    Next i\
\
   'Section 6 going after volume change in column N. \
\
\
\'91 ****REQUIREMENTS*******    \
\'91Greatest Total Volume (5 points)\
    \
    Dim largestVolume As Double\
    Dim currentVolume As Double\
    \
    currentVolme = 0\
    \
    For i = 2 To 3001 ' created a 2nd column length checker\
        currentVolume = Range("L" & i).Value ' k is percent change column\
        currentStock = Range("I" & i).Value ' same  assumption here\
        If currentVolume > largestVolume Then\
            largestVolume = currentVolume\
           Range("Q4").Value = largestVolume\
           Range("P4").Value = currentStock\
\
        End If\
    Next i\
\
'i = 0\
'lastRow = 0\
'lastRow2 = 0\
'currentValue = 0\
\
\
End Sub\
\
\
\
}