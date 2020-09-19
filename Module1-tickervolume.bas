Attribute VB_Name = "Module1"
Sub tickervolumes()
'Define Variable for Tickers
Dim TickerName As String

' Set initial variable for holding the total per ticker
Dim TotalVolume As Variant
TotalVolume = 0

'Keep track of the locations for each ticker in summary
Dim SummaryTable As Integer
SummaryTable = 2

' Loop through all the tickers
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row

' Check if we are still within the same ticker value
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

    ' Set the ticker name
    TickerName = Cells(i, 1).Value
    
    ' Add to total volume
    TotalVolume = TotalVolume + Cells(i, 7).Value
    
    ' Print ticker value in summary table
    Range("I" & SummaryTable).Value = TickerName
    
    ' Print total volume in summary table
    Range("L" & SummaryTable).Value = TotalVolume
    
    ' Add one to the summary row to go to next row
    SummaryTable = SummaryTable + 1
    
    ' Reset Value of summary table
    TotalVolume = 0

'If the immediate row is the same value then..
Else
    ' Add to the Volume total
    TotalVolume = TotalVolume + Cells(i, 7).Value
    End If
    
    Next i

End Sub
