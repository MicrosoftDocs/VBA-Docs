---
title: Application.InchesToPoints method (Excel)
keywords: vbaxl10.chm133148
f1_keywords:
- vbaxl10.chm133148
ms.prod: excel
api_name:
- Excel.Application.InchesToPoints
ms.assetid: 7689eae4-f533-32e3-d431-4873029a8bc1
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.InchesToPoints method (Excel)

Converts a measurement from inches to points.


## Syntax

_expression_.**InchesToPoints** (_Inches_)

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Inches_|Required| **Double**|Specifies the inch value to be converted to points.|

## Return value

Double


## Example

This example sets the left margin of Sheet1 to 2.5 inches.

```vb
Worksheets("Sheet1").PageSetup.LeftMargin = _ 
 Application.InchesToPoints(2.5)
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]