---
title: Worksheet.Cells property (Excel)
keywords: vbaxl10.chm175080
f1_keywords:
- vbaxl10.chm175080
ms.prod: excel
api_name:
- Excel.Worksheet.Cells
ms.assetid: 19c14e41-7d8e-b56f-fd60-717df64edee8
ms.date: 05/30/2019
localization_priority: Normal
---


# Worksheet.Cells property (Excel)

Returns a **[Range](Excel.Range(object).md)** object that represents all the cells on the worksheet (not just the cells that are currently in use).


## Syntax

_expression_.**Cells**

_expression_ A variable that represents a **[Worksheet](Excel.Worksheet.md)** object.


## Remarks

Because the default member of **Range** forwards calls with parameters to the **[Item](Excel.Range.Item.md)** property, you can specify the row and column index immediately after the **Cells** keyword instead of an explicit call to **[Item](Excel.Range.Item.md)**.

Using this property without an object qualifier returns a **Range** object that represents all the cells on the active worksheet.


## Example

This example sets the font size for cell C5 on Sheet1 of the active workbook to 14 points.

```vb
Worksheets("Sheet1").Cells(5, 3).Font.Size = 14
```

<br/>

This example clears the formula in cell one on Sheet1 of the active workbook.

```vb
Worksheets("Sheet1").Cells(1).ClearContents
```

<br/>

This example sets the font and font size for every cell on Sheet1 to 8-point Arial.

```vb
With Worksheets("Sheet1").Cells.Font 
    .Name = "Arial" 
    .Size = 8 
End With
```

<br/>

This example toggles a sort between ascending and descending order when you double-click any cell in the data range. The data is sorted based on the column of the cell that is double-clicked.

```vb
Option Explicit
Public blnToggle As Boolean

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Dim LastColumn As Long, keyColumn As Long, LastRow As Long
    Dim SortRange As Range
    LastColumn = Cells.Find(What:="*", After:=Range("A1"), SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    
    keyColumn = Target.Column
    
    If keyColumn <= LastColumn Then
    
        Application.ScreenUpdating = False
        Cancel = True
        LastRow = Cells(Rows.Count, keyColumn).End(xlUp).Row
        Set SortRange = Target.CurrentRegion
        
        blnToggle = Not blnToggle
        If blnToggle = True Then
            SortRange.Sort Key1:=Cells(2, keyColumn), Order1:=xlAscending, Header:=xlYes
        Else
            SortRange.Sort Key1:=Cells(2, keyColumn), Order1:=xlDescending, Header:=xlYes
        End If
    
        Set SortRange = Nothing
        Application.ScreenUpdating = True
        
    End If
End Sub
```

<br/>

This example looks through column C of the active sheet, and for every cell that has a comment, it puts the comment text into column D and deletes the comment from column C.

```vb
Public Sub SplitCommentsOnActiveSheet()
   'Set up your variables
   Dim cmt As Comment
   Dim rowIndex As Integer
   
   'Go through all the cells in Column C, and check to see if the cell has a comment.
   For rowIndex = 1 To WorksheetFunction.CountA(Columns(3))
      Set cmt = Cells(rowIndex, 3).Comment
      If Not cmt Is Nothing Then
      
         'If there is a comment, paste the comment text into column D and delete the original comment.
         Cells(rowIndex, 4) = Cells(rowIndex, 3).Comment.Text
         Cells(rowIndex, 3).Comment.Delete
      End If
   Next
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
