---
title: Range.AddIndent Property (Excel)
keywords: vbaxl10.chm144075
f1_keywords:
- vbaxl10.chm144075
ms.prod: excel
api_name:
- Excel.Range.AddIndent
ms.assetid: 47cfb2a4-9050-354f-08f6-e86f0164be02
ms.date: 06/08/2017
---


# Range.AddIndent Property (Excel)

Returns or sets a  **Variant** value that indicates if text is automatically indented when the text alignment in a cell is set to equal distribution (either horizontally or vertically.)


## Syntax

 _expression_ . **AddIndent**

 _expression_ A variable that represents a **Range** object.


## Remarks

Set the value of this property to  **True** to autmatically indent text when the text alignment in the cell is set, either horizontally or vertically, to equal distribution.

To set text alignment to equal distribution, you can set the  **[VerticalAlignment](Excel.Range.VerticalAlignment.md)** property to **xlVAlignDistributed** when the value of the **[Orientation](Excel.Range.Orientation.md)** property is **xlVertical** , and you can set the **[HorizontalAlignment](Excel.Range.HorizontalAlignment.md)** property to **xlHAlignDistributed** when the value of the **Orientation** property is **xlHorizontal** .


## Example

This example sets the horizontal alignment for text in cell A1 on Sheet1 to equal distribution and then indents the text.


```vb
With Worksheets("Sheet1").Range("A1") 
 .HorizontalAlignment = xlHAlignDistributed 
 .AddIndent = True 
End With
```


## See also


#### Concepts


[Range Object](Excel.Range(objec).md)

