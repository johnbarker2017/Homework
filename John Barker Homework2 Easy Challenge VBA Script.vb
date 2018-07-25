Sub HomeworkEasyChallenge()

    
Dim ws As Worksheet
Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet 'remember which worksheet is active in the beginning
For Each ws In ThisWorkbook.Worksheets
    ws.Activate
    Columns("A").Copy
    Columns("I").Insert
    Columns("G").Copy
    Columns("J").Insert
    'Rename Columns
 Range("I1").Select
    ActiveCell.FormulaR1C1 = "Ticker"
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "Total Stock Volume"

    'Fit Column Width to Contents
  For Each sht In ThisWorkbook.Worksheets
  sht.Cells.EntireColumn.AutoFit
  ActiveWorkbook.Save
  Next
 ' Next sht

    
 Next

starting_ws.Activate 'activate the worksheet that was originally active
    

    
End Sub 


