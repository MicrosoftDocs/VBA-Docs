---
title: Range.FunctionWizard method (Excel)
keywords: vbaxl10.chm144139
f1_keywords:
- vbaxl10.chm144139
ms.prod: excel
api_name:
- Excel.Range.FunctionWizard
ms.assetid: a9a0c765-4903-4969-8f09-c8f051213a96
ms.date: 05/10/2019
localization_priority: Normal
---


# Range.FunctionWizard method (Excel)

Starts the Function Wizard for the upper-left cell of the range.


## Syntax

_expression_.**FunctionWizard**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Return value

Variant


## Example

This example starts the Function Wizard for the active cell on Sheet1.

```vb
Worksheets("Sheet1").Activate 
ActiveCell.FunctionWizard
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]