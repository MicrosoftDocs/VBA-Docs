---
title: Range.NumberFormat property (Excel)
keywords: vbaxl10.chm144167
f1_keywords:
- vbaxl10.chm144167
ms.prod: excel
api_name:
- Excel.Range.NumberFormat
ms.assetid: 351247d2-e4b9-64a0-6dbe-0df535fa701c
ms.date: 05/11/2019
localization_priority: Priority
---

# Range.NumberFormat property (Excel)

Returns or sets a **Variant** value that represents the format code for the object.

## Syntax

_expression_.**NumberFormat**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.

## Remarks

This property returns **Null** if all cells in the specified range don't have the same number format.

The format code is the same string as the **Format Codes** option in the **Format Cells** dialog box. The **Format** function uses different format code strings than do the **NumberFormat** and **[NumberFormatLocal](Excel.Range.NumberFormatLocal.md)** properties.

For more information, see [Number format codes (Microsoft Support)](https://support.office.com/article/number-format-codes-5026bbd6-04bc-48cd-bf33-80f18b4eae68).

## Example

These examples set the number format for cell A17, row one, and column C (respectively) on Sheet1.

```vb
Worksheets("Sheet1").Range("A17").NumberFormat = "General" 
Worksheets("Sheet1").Rows(1).NumberFormat = "hh:mm:ss" 
Worksheets("Sheet1").Columns("C"). _ 
 NumberFormat = "$#,##0.00_);[Red]($#,##0.00)"
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
