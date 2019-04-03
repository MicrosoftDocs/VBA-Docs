---
title: Application.Calculation property (Excel)
keywords: vbaxl10.chm133084
f1_keywords:
- vbaxl10.chm133084
ms.prod: excel
api_name:
- Excel.Application.Calculation
ms.assetid: 5ae7f2dd-e79a-a4ee-f701-2fff1b77f499
ms.date: 04/04/2019
localization_priority: Normal
---


# Application.Calculation property (Excel)

Returns or sets an **[XlCalculation](Excel.XlCalculation.md)** value that represents the calculation mode.


## Syntax

_expression_.**Calculation**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Remarks

For OLAP data sources, this property can only return or be set to **xlNormal**.


## Example

This example causes Microsoft Excel to calculate workbooks before they are saved to disk.

```vb
Application.Calculation = xlCalculationManual 
Application.CalculateBeforeSave = True
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
