---
title: Validation.IgnoreBlank property (Excel)
keywords: vbaxl10.chm532075
f1_keywords:
- vbaxl10.chm532075
ms.prod: excel
api_name:
- Excel.Validation.IgnoreBlank
ms.assetid: 91913061-9cc7-8e96-11c3-67d7b84e2e25
ms.date: 05/18/2019
localization_priority: Normal
---


# Validation.IgnoreBlank property (Excel)

**True** if blank values are permitted by the range data validation. Read/write **Boolean**.


## Syntax

_expression_.**IgnoreBlank**

_expression_ A variable that represents a **[Validation](Excel.Validation.md)** object.


## Remarks

If the **IgnoreBlank** property is **True**, cell data is considered valid if the cell is blank, or if a cell referenced by either the **MinVal** or **MaxVal** property is blank.


## Example

This example causes data validation for cell E5 to allow blank values.

```vb
Range("e5").Validation.IgnoreBlank = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]