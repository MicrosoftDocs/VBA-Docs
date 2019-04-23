---
title: ChartObjects.Placement property (Excel)
keywords: vbaxl10.chm497086
f1_keywords:
- vbaxl10.chm497086
ms.prod: excel
api_name:
- Excel.ChartObjects.Placement
ms.assetid: 954e98e5-8b88-6918-3cbd-f8e982c0a47e
ms.date: 04/20/2019
localization_priority: Normal
---


# ChartObjects.Placement property (Excel)

Returns or sets a **Variant** value, containing an **[XlPlacement](Excel.XlPlacement.md)** constant, that represents the way the objects are attached to the cells below them.


## Syntax

_expression_.**Placement**

_expression_ A variable that represents a **[ChartObjects](Excel.ChartObjects.md)** object.


## Example

This example sets the objects on Sheet1 to be free-floating (they neither move nor are they sized with underlying cells).

```vb
Worksheets("Sheet1").ChartObjects.Placement = xlFreeFloating
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]