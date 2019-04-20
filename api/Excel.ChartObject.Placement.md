---
title: ChartObject.Placement property (Excel)
keywords: vbaxl10.chm494088
f1_keywords:
- vbaxl10.chm494088
ms.prod: excel
api_name:
- Excel.ChartObject.Placement
ms.assetid: 61369038-c3ab-531f-93c0-b8bdfe3c07dd
ms.date: 04/20/2019
localization_priority: Normal
---


# ChartObject.Placement property (Excel)

Returns or sets a **Variant** value, containing an **[XlPlacement](Excel.XlPlacement.md)** constant, that represents the way the object is attached to the cells below it.


## Syntax

_expression_.**Placement**

_expression_ A variable that represents a **[ChartObject](Excel.ChartObject.md)** object.


## Example

This example sets embedded chart one on Sheet1 to be free-floating (it neither moves nor is sized with its underlying cells).

```vb
Worksheets("Sheet1").ChartObjects(1).Placement = xlFreeFloating
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]