---
title: Range.Cells property (Excel)
keywords: vbaxl10.chm144091
f1_keywords:
- vbaxl10.chm144091
ms.prod: excel
api_name:
- Excel.Range.Cells
ms.assetid: 32a6ecc7-2366-2cec-1feb-0966241a435d
ms.date: 08/14/2019
localization_priority: Priority
---


# Range.Cells property (Excel)

Returns a **Range** object that represents the cells in the specified range.

[!include[Add-ins note](../includes/addinsnote.md)]


## Syntax

_expression_.**Cells**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Remarks

Because the **[Item](Excel.Range.Item.md)** property is the default property for the **Range** object, you can specify the row and column index immediately after the **Cells** keyword. For more information, see the **Item** property and the examples for this topic.

Using this property without an object qualifier returns a **Range** object that represents all the cells on the active worksheet.


## Example

This example sets the font style for cells A1:C5 on Sheet1 to italic.

```vb
Worksheets("Sheet1").Activate 
Range(Cells(1, 1), Cells(5, 3)).Font.Italic = True
```

<br/>

This example scans a column of data named _myRange_. If a cell has the same value as the cell immediately preceding it, the example displays the address of the cell that contains the duplicate data.

```vb
Set r = Range("myRange") 
For n = 2 To r.Rows.Count 
    If r.Cells(n-1, 1) = r.Cells(n, 1) Then 
        MsgBox "Duplicate data in " & r.Cells(n, 1).Address 
    End If 
Next n
```

<br/>

This example looks through column C, and for every cell that has a comment, it puts the comment text into column D and deletes the comment from column C.

```vb
Sub SplitComments()
   'Set up your variables
   Dim cmt As Comment
   Dim iRow As Integer
   
   'Go through all the cells in Column C, and check to see if the cell has a comment.
   For iRow = 1 To WorksheetFunction.CountA(Columns(3))
      Set cmt = Cells(iRow, 3).Comment
      If Not cmt Is Nothing Then
      
         'If there is a comment, paste the comment text into column D and delete the original comment.
         Cells(iRow, 4) = Cells(iRow, 3).Comment.Text
         Cells(iRow, 3).Comment.Delete
      End If
   Next iRow
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
