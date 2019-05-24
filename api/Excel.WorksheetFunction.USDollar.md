---
title: WorksheetFunction.USDollar method (Excel)
keywords: vbaxl10.chm137153
f1_keywords:
- vbaxl10.chm137153
ms.prod: excel
api_name:
- Excel.WorksheetFunction.USDollar
ms.assetid: d09c7356-d6c1-0290-5ed8-ed9c3732a21b
ms.date: 05/25/2019
localization_priority: Normal
---


# WorksheetFunction.USDollar method (Excel)

Converts a number to text format and applies a currency symbol. The name of the method (and the symbol that it applies) depends upon the language settings.


## Syntax

_expression_.**USDollar** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|A reference to a cell containing a number, or a formula that evaluates to a number.|
| _Arg2_|Required| **Double**|The number of digits to the right of the decimal point. If _Arg2_ is negative, the number is rounded to the left of the decimal point. If you omit decimals, it is assumed to be 2.|

## Return value

**String**


## Remarks

The **USDollar** method converts a number to text by using currency format, with the decimals rounded to the specified place. The format used is `$#,##0.00_);($#,##0.00)`.

The major difference between formatting a cell that contains a number with the **Format Cells** command and formatting a number directly with the **[Dollar](excel.worksheetfunction.dollar.md)** method is that **Dollar** converts its result to text. A number formatted with the **Format Cells** command is still a number. You can continue to use numbers formatted with **Dollar** in formulas, because Excel converts numbers entered as text values to numbers when it calculates.


## Example

The following example displays the first number in a currency format, two digits to the right of the decimal point ($1,234.57).

```vb
=DOLLAR(A2, 2)
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]