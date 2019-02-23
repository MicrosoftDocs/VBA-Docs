---
title: Range.RowHeight property (Excel)
keywords: vbaxl10.chm144190
f1_keywords:
- vbaxl10.chm144190
ms.prod: excel
api_name:
- Excel.Range.RowHeight
ms.assetid: 103c7209-9a4f-8f9c-7bdc-3013113867a5
ms.date: 09/05/2018
localization_priority: Priority
---


# Range.RowHeight property (Excel)

Returns or sets the height of the first row in the range specified, measured in points. Read/write **Double**.


## Syntax

_expression_. `RowHeight`

_expression_ A variable that represents a [Range](excel.range-graph-property.md) object.


## Remarks

**RowHeight** property sets the height for all rows in a range of cells.

Use the **[AutoFit](Excel.Range.AutoFit.md)** method to set row heights based on the contents of cells.

> [!NOTE]
> If a merged cell is in the range, **RowHeight** returns **Null** for varied row heights. Use the **[Height](Excel.Range.Height.md)** property to return the total height of a range of cells.

> [!NOTE]
> When a range contains rows of different heights, **RowHeight** might return the height of the first row or might return **Null**.


## Example

This example doubles the height of row one on Sheet1.


```vb
With Worksheets("Sheet1").Rows(1) 
 .RowHeight = .RowHeight * 2 
End With
```


## See also

- [Slicer.RowHeight property](Excel.Slicer.RowHeight.md)
- [Range object](Excel.Range(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]