---
title: Range.Columns property (Excel)
keywords: vbaxl10.chm144101
f1_keywords:
- vbaxl10.chm144101
ms.prod: excel
api_name:
- Excel.Range.Columns
ms.assetid: a1a23288-e911-909d-0bc0-48bdce2ccbac
ms.date: 05/10/2019
localization_priority: Normal
---


# Range.Columns property (Excel)

Returns a **Range** object that represents the columns in the specified range.


## Syntax

_expression_.**Columns**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Remarks

To return a single column, use the **[Item](Excel.Range.Item.md)** property or equivalently include an index in parentheses. For example, both `Selection.Columns(1)` and `Selection.Columns.Item(1)` return the first column of the selection.

When applied to a **Range** object that is a multiple-area selection, this property returns columns from only the first area of the range. For example, if the **Range** object has two areas—A1:B2 and C3:D4—`Selection.Columns.Count` returns 2, not 4. To use this property on a range that may contain a multiple-area selection, test `Areas.Count` to determine whether the range contains more than one area. If it does, loop over each area in the range.

The returned range might be outside the specified range. For example, `Range("A1:B2").Columns(5).Select` returns cells E1:E2.

If a letter is used as an index, it is equivalent to a number. For example, `Range("B1:C10").Columns("B").Select` returns cells C1:C10, not cells B1:B10. In the example, "B" is equivalent to 2.

Using the **Columns** property without an object qualifier is equivalent to using `ActiveSheet.Columns`. For more information, see the **[Worksheet.Columns](excel.worksheet.columns.md)** property.

## Example

This example sets the value of every cell in column one in the range named _myRange_ to 0 (zero).

```vb
Range("myRange").Columns(1).Value = 0
```

<br/>

This example displays the number of columns in the selection on Sheet1. If more than one area is selected, the example loops through each area.

```vb
Public Sub ShowNumberOfColumnsInSheet1Selection
   Worksheets("Sheet1").Activate 
   
   Dim selectedRange As Excel.Range
   Set selectedRange = Selection
   
   Dim areaCount As Long
   areaCount = Selection.Areas.Count 
   
   If areaCount <= 1 Then 
      MsgBox "The selection contains " & _ 
             Selection.Columns.Count & " columns." 
   Else 
      Dim areaIndex As Long
      areaIndex = 1 
      For Each area In Selection.Areas 
         MsgBox "Area " & areaIndex & " of the selection contains " & _ 
                area.Columns.Count & " columns." 
         areaIndex = areaIndex + 1 
      Next 
   End If
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
