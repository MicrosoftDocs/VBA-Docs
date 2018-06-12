---
title: CellFormat.AddIndent Property (Excel)
keywords: vbaxl10.chm676078
f1_keywords:
- vbaxl10.chm676078
ms.prod: excel
api_name:
- Excel.CellFormat.AddIndent
ms.assetid: 7f38c3d8-ccea-fc6c-a171-d028fe30080d
ms.date: 06/08/2017
---


# CellFormat.AddIndent Property (Excel)

Returns or sets a  **Variant** value that indicates if text is automatically indented when the text alignment in a cell is set to equal distribution (either horizontally or vertically.)


## Syntax

 _expression_ . **AddIndent**

 _expression_ A variable that represents a **CellFormat** object.


## Remarks

Set the value of this property to  **True** to autmatically indent text when the text alignment in the cell is set, either horizontally or vertically, to equal distribution.

To set text alignment to equal distribution, you can set the  **[VerticalAlignment](Excel.Range.VerticalAlignment.md)** property to **xlVAlignDistributed** when the value of the **[Orientation](Excel.Range.Orientation.md)** property is **xlVertical** , and you can set the **[HorizontalAlignment](Excel.Range.HorizontalAlignment.md)** property to **xlHAlignDistributed** when the value of the **Orientation** property is **xlHorizontal** .


## See also


#### Concepts


[CellFormat Object](Excel.CellFormat.md)

