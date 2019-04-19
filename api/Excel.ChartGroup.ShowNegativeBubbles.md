---
title: ChartGroup.ShowNegativeBubbles property (Excel)
keywords: vbaxl10.chm568096
f1_keywords:
- vbaxl10.chm568096
ms.prod: excel
api_name:
- Excel.ChartGroup.ShowNegativeBubbles
ms.assetid: 1f1288d5-71c5-f5da-583c-584db90c6c33
ms.date: 04/20/2019
localization_priority: Normal
---


# ChartGroup.ShowNegativeBubbles property (Excel)

**True** if negative bubbles are shown for the chart group. Valid only for bubble charts. Read/write **Boolean**.


## Syntax

_expression_.**ShowNegativeBubbles**

_expression_ A variable that represents a **[ChartGroup](Excel.ChartGroup(object).md)** object.


## Example


```vb
Worksheets(1).ChartObjects(1).Chart _ 
 .ChartGroups(1).ShowNegativeBubbles = True
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]