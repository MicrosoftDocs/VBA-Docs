---
title: CellFormat.VerticalAlignment property (Excel)
keywords: vbaxl10.chm676081
f1_keywords:
- vbaxl10.chm676081
ms.prod: excel
api_name:
- Excel.CellFormat.VerticalAlignment
ms.assetid: c901dff3-3f0a-1f54-250e-c03b9e32c819
ms.date: 04/16/2019
localization_priority: Normal
---


# CellFormat.VerticalAlignment property (Excel)

Returns or sets a **Variant** value that represents the vertical alignment of the specified object.


## Syntax

_expression_.**VerticalAlignment**

_expression_ A variable that represents a **[CellFormat](Excel.CellFormat.md)** object.


## Remarks

TThe value of this property can be set to one of the **[XlVAlign](excel.xlvalign.md)** constants.

## Example

This example sets the height of row 2 on Sheet1 to twice the standard height, and then centers the contents of the row vertically.

```vb
Worksheets("Sheet1").Rows(2).RowHeight = _ 
 2 * Worksheets("Sheet1").StandardHeight 
Worksheets("Sheet1").Rows(2).VerticalAlignment = xlVAlignCenter 

```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
