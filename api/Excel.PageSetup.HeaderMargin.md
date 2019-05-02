---
title: PageSetup.HeaderMargin property (Excel)
keywords: vbaxl10.chm473085
f1_keywords:
- vbaxl10.chm473085
ms.prod: excel
api_name:
- Excel.PageSetup.HeaderMargin
ms.assetid: c22feaf6-c9f5-f285-a8f6-75753a1e9cff
ms.date: 05/03/2019
localization_priority: Normal
---


# PageSetup.HeaderMargin property (Excel)

Returns or sets the distance from the top of the page to the header, in [points](../language/glossary/vbe-glossary.md#point). Read/write **Double**.


## Syntax

_expression_.**HeaderMargin**

_expression_ A variable that represents a **[PageSetup](Excel.PageSetup.md)** object.


## Remarks

Margins are set or returned in points. Use either the **[InchesToPoints](Excel.Application.InchesToPoints.md)** method or the **[CentimetersToPoints](Excel.Application.CentimetersToPoints.md)** method to do the conversion.


## Example

This example sets the header margin of Sheet1 to 0.5 inch.

```vb
Worksheets("Sheet1").PageSetup.HeaderMargin = _ 
 Application.InchesToPoints(0.5)
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]