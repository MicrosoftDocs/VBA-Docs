---
title: Range.Validation property (Excel)
keywords: vbaxl10.chm144215
f1_keywords:
- vbaxl10.chm144215
ms.prod: excel
api_name:
- Excel.Range.Validation
ms.assetid: d1cad7e6-bbfa-e280-33e7-048733efc0bc
ms.date: 05/11/2019
localization_priority: Normal
---


# Range.Validation property (Excel)

Returns the **[Validation](Excel.Validation.md)** object that represents data validation for the specified range. Read-only.


## Syntax

_expression_.**Validation**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Example

This example causes data validation for cell E5 to allow blank values.

```vb
Range("e5").Validation.IgnoreBlank = True
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
