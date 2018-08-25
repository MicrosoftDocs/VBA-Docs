---
title: Range.ColumnWidth Property (Excel)
keywords: vbaxl10.chm144102
f1_keywords:
- vbaxl10.chm144102
ms.prod: excel
api_name:
- Excel.Range.ColumnWidth
ms.assetid: a6364bb1-2e3d-07d6-20e4-c9fa8f7c5ad3
ms.date: 08/24/2018
---


# Range.ColumnWidth Property (Excel)

Returns or sets the width of all columns in the specified range. Read/write **Variant** .


## Syntax

 _expression_. `ColumnWidth`

 _expression_ A variable that represents a [Range](./Excel.Range(Graph property).md) object.


## Remarks

One unit of column width is equal to the width of one character in the Normal style. For proportional fonts, the width of the character 0 (zero) is used.

Use the **[AutoFit](range-autofit-method-excel.md)** method to set column widths based on the contents of cells.

Use the **[Width](Excel.Range.Width.md)** property to return the width of a column in points.

If all columns in the range have the same width, the **ColumnWidth** property returns the width. If columns in the range have different widths, this property returns **null** .


## Example

The following example doubles the width of column A on Sheet1.


```vb
With Worksheets("Sheet1").Columns("A") 
 .ColumnWidth = .ColumnWidth * 2 
End With
```


## See also
[Range Object](Excel.Range(object).md)
[Range.Width Property](range-width-property-excel.md)
[Range.AutoFit Method](range-autofit-method-excel.md)

