---
title: CellFormat.AddIndent property (Excel)
keywords: vbaxl10.chm676078
f1_keywords:
- vbaxl10.chm676078
ms.prod: excel
api_name:
- Excel.CellFormat.AddIndent
ms.assetid: 7f38c3d8-ccea-fc6c-a171-d028fe30080d
ms.date: 04/16/2019
localization_priority: Normal
---


# CellFormat.AddIndent property (Excel)

Returns or sets a **Variant** value that indicates if text is automatically indented when the text alignment in a cell is set to equal distribution (either horizontally or vertically).


## Syntax

_expression_.**AddIndent**

_expression_ A variable that represents a **[CellFormat](Excel.CellFormat.md)** object.


## Remarks

Set the value of this property to **True** to automatically indent text when the text alignment in the cell is set, either horizontally or vertically, to equal distribution.

To set text alignment to equal distribution, you can set the **[VerticalAlignment](Excel.Range.VerticalAlignment.md)** property to **xlVAlignDistributed** when the value of the **[Orientation](Excel.Range.Orientation.md)** property is **xlVertical**, and you can set the **[HorizontalAlignment](Excel.Range.HorizontalAlignment.md)** property to **xlHAlignDistributed** when the value of the **Orientation** property is **xlHorizontal**.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]