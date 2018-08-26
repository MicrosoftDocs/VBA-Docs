---
title: Range.ColumnWidth Property (Excel)
keywords: vbaxl10.chm144102
f1_keywords:
- vbaxl10.chm144102
ms.prod: excel
api_name:
- Excel.Range.ColumnWidth
ms.assetid: a6364bb1-2e3d-07d6-20e4-c9fa8f7c5ad3
<<<<<<< HEAD
ms.date: 06/08/2017
=======
ms.date: 08/24/2018
>>>>>>> master
---


# Range.ColumnWidth Property (Excel)

<<<<<<< HEAD
Returns or sets the width of all columns in the specified range. Read/write  **Variant** .
=======
Returns or sets the width of all columns in the specified range. Read/write **Variant** .
>>>>>>> master


## Syntax

 _expression_. `ColumnWidth`

<<<<<<< HEAD
 _expression_ A variable that represents a [Range](./Excel.Range(Graph property).md) object.
=======
 _expression_ A variable that represents a [Range](Excel.Range(Graph property).md) object.
>>>>>>> master


## Remarks

One unit of column width is equal to the width of one character in the Normal style. For proportional fonts, the width of the character 0 (zero) is used.

<<<<<<< HEAD
Use the  **[Width](Excel.Range.Width.md)** property to return the width of a column in points.

If all columns in the range have the same width, the  **ColumnWidth** property returns the width. If columns in the range have different widths, this property returns **null** .
=======
Use the **[AutoFit](Excel.Autofit.md)** method to set column widths based on the contents of cells.

Use the **[Width](Excel.Width.md)** property to return the width of a column in points.

If all columns in the range have the same width, the **ColumnWidth** property returns the width. If columns in the range have different widths, this property returns **null** .
>>>>>>> master


## Example

<<<<<<< HEAD
This example doubles the width of column A on Sheet1.
=======
The following example doubles the width of column A on Sheet1.
>>>>>>> master


```vb
With Worksheets("Sheet1").Columns("A") 
 .ColumnWidth = .ColumnWidth * 2 
End With
```


## See also

<<<<<<< HEAD

[Range Object](Excel.Range(object).md)
=======
- [Range Object](Excel.Range(object).md)

>>>>>>> master

