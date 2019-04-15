---
title: CellFormat.WrapText property (Excel)
keywords: vbaxl10.chm676084
f1_keywords:
- vbaxl10.chm676084
ms.prod: excel
api_name:
- Excel.CellFormat.WrapText
ms.assetid: 92d7920c-51e2-f949-60ee-d11595c191bb
ms.date: 04/16/2019
localization_priority: Normal
---


# CellFormat.WrapText property (Excel)

Returns or sets a **Variant** value that indicates if Microsoft Excel wraps the text in the object.


## Syntax

_expression_.**WrapText**

_expression_ A variable that represents a **[CellFormat](Excel.CellFormat.md)** object.


## Remarks

This property returns **True** if text is wrapped in all cells within the specified range, **False** if text is not wrapped in all cells within the specified range, or **Null** if the specified range contains some cells that wrap text and other cells that don't.

Microsoft Excel will change the row height of the range, if necessary, to accommodate the text in the range.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]