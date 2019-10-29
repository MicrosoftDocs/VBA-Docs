---
title: Range.Rows property (Excel)
keywords: vbaxl10.chm144191
f1_keywords:
- vbaxl10.chm144191
ms.prod: excel
api_name:
- Excel.Range.Rows
ms.assetid: 2b0541f1-119d-8535-8418-ff9482353ec1
ms.date: 05/11/2019
localization_priority: Normal
---


# Range.Rows property (Excel)

Returns a **Range** object that represents the rows in the specified range.


## Syntax

_expression_.**Rows**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Remarks

To return a single row, use the **[Item](Excel.Range.Item.md)** property or equivalently include an index in parentheses. For example, both `Selection.Rows(1)` and `Selection.Rows.Item(1)` return the first row of the selection.

When applied to a **Range** object that is a multiple selection, this property returns rows from only the first area of the range. For example, if the **Range** object `someRange` has two areas—A1:B2 and C3:D4—,`someRange.Rows.Count` returns 2, not 4. To use this property on a range that may contain a multiple selection, test **Areas.Count** to determine whether the range is a multiple selection. If it is, loop over each area in the range, as shown in the third example.

The returned range might be outside the specified range. For example, `Range("A1:B2").Rows(5)` returns cells A5:B5. For more information, see the **[Item](Excel.Range.Item.md)** property.

Using the **Rows** property without an object qualifier is equivalent to using **ActiveSheet.Rows**. For more information, see the **[Worksheet.Rows](excel.worksheet.rows.md)** property.


## Example

This example deletes the range B5:Z5 on Sheet1 of the active workbook.

```vb
Worksheets("Sheet1").Range("B2:Z44").Rows(3).Delete
```

<br/>

This example deletes rows in the current region on worksheet one of the active workbook where the value of cell one in the row is the same as the value of cell one in the previous row.

```vb
For Each rw In Worksheets(1).Cells(1, 1).CurrentRegion.Rows
   this = rw.Cells(1, 1).Value 
   If this = last Then rw.Delete 
   last = this 
Next
```

<br/>

This example displays the number of rows in the selection on Sheet1. If more than one area is selected, the example loops through each area.

```vb
Public Sub ShowNumberOfRowsInSheet1Selection
   Worksheets("Sheet1").Activate 
   
   Dim selectedRange As Excel.Range
   Set selectedRange = Selection
   
   Dim areaCount As Long
   areaCount = Selection.Areas.Count 
   
   If areaCount <= 1 Then 
      MsgBox "The selection contains " & _ 
             Selection.Rows.Count & " rows." 
   Else 
      Dim areaIndex As Long
      areaIndex = 1 
      For Each area In Selection.Areas 
         MsgBox "Area " & areaIndex & " of the selection contains " & _ 
                area.Rows.Count & " rows." 
         areaIndex = areaIndex + 1 
      Next 
   End If
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
