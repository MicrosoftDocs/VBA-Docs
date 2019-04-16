---
title: CellFormat.NumberFormat property (Excel)
keywords: vbaxl10.chm676076
f1_keywords:
- vbaxl10.chm676076
ms.prod: excel
api_name:
- Excel.CellFormat.NumberFormat
ms.assetid: 55133c7e-7d55-a2a9-0a76-9bd630a59cc4
ms.date: 04/16/2019
localization_priority: Normal
---


# CellFormat.NumberFormat property (Excel)

Returns or sets a **Variant** value that represents the format code for the object.


## Syntax

_expression_.**NumberFormat**

_expression_ A variable that represents a **[CellFormat](Excel.CellFormat.md)** object.


## Remarks

This property returns **Null** if all cells in the specified range don't have the same number format.

The format code is the same string as the **Format Codes** option in the **Format Cells** dialog box. The **Format** function uses different format code strings than do the **NumberFormat** and **[NumberFormatLocal](Excel.CellFormat.NumberFormatLocal.md)** properties.

## Example

The following example is for the **NumberFormat** function. 

```vba

Range("A1:A5").Value = 12345
Range("A1:A5").NumberFormat = "0.00"
Range("A1:A5").NumberFormat = "General"

```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]