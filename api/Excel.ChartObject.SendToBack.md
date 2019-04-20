---
title: ChartObject.SendToBack method (Excel)
keywords: vbaxl10.chm494091
f1_keywords:
- vbaxl10.chm494091
ms.prod: excel
api_name:
- Excel.ChartObject.SendToBack
ms.assetid: a8f0f721-15ba-662f-ac17-0ac1657e3413
ms.date: 04/20/2019
localization_priority: Normal
---


# ChartObject.SendToBack method (Excel)

Sends the object to the back of the z-order.


## Syntax

_expression_.**SendToBack**

_expression_ A variable that represents a **[ChartObject](Excel.ChartObject.md)** object.


## Return value

Variant


## Example

This example sends embedded chart one on Sheet1 to the back of the z-order.

```vb
Worksheets("Sheet1").ChartObjects(1).SendToBack
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]