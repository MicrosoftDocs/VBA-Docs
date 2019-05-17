---
title: Style.AddIndent property (Excel)
keywords: vbaxl10.chm177073
f1_keywords:
- vbaxl10.chm177073
ms.prod: excel
api_name:
- Excel.Style.AddIndent
ms.assetid: 76b9c820-8c94-3cf6-7267-6d2710f07b74
ms.date: 05/16/2019
localization_priority: Normal
---


# Style.AddIndent property (Excel)

Returns or sets a **Boolean** value that indicates if text is automatically indented when the text alignment in a cell is set to equal distribution (either horizontally or vertically).


## Syntax

_expression_.**AddIndent**

_expression_ A variable that represents a **[Style](Excel.Style.md)** object.


## Remarks

Set the value of this property to **True** to automatically indent text when the text alignment in the cell is set, either horizontally or vertically, to equal distribution.

To set text alignment to equal distribution, you can set the **[VerticalAlignment](Excel.Range.VerticalAlignment.md)** property of the **Range** object to **xlVAlignDistributed** when the value of the **[Orientation](Excel.Range.Orientation.md)** property is **xlVertical**, and you can set the **[HorizontalAlignment](Excel.Range.HorizontalAlignment.md)** property to **xlHAlignDistributed** when the value of the **Orientation** property is **xlHorizontal**.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]