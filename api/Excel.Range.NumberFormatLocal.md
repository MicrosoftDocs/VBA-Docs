---
title: Range.NumberFormatLocal property (Excel)
keywords: vbaxl10.chm144168
f1_keywords:
- vbaxl10.chm144168
ms.prod: excel
api_name:
- Excel.Range.NumberFormatLocal
ms.assetid: e34e6f52-9279-7961-adfa-4aa84c44937a
ms.date: 05/11/2019
localization_priority: Normal
---

# Range.NumberFormatLocal property (Excel)

Returns or sets a **Variant** value that represents the format code for the object as a string in the language of the user.

## Syntax

_expression_.**NumberFormatLocal**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.

## Remarks

The **Format** function uses different format code strings than do the **[NumberFormat](Excel.Range.NumberFormat.md)** and **NumberFormatLocal** properties.

For more information, see [Number format codes (Microsoft Support)](https://support.office.com/article/number-format-codes-5026bbd6-04bc-48cd-bf33-80f18b4eae68).


## Example

This example displays the number format for cell A1 on Sheet1 in the language of the user.

```vb
MsgBox "The number format for cell A1 is " & _ 
 Worksheets("Sheet1").Range("A1").NumberFormatLocal
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
