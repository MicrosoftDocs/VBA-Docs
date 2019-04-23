---
title: ShowNegativeBubbles property (Excel Graph)
keywords: vbagr10.chm67190
f1_keywords:
- vbagr10.chm67190
ms.prod: excel
api_name:
- Excel.ShowNegativeBubbles
ms.assetid: 1ef1b415-8e89-a57d-249c-db7e85086d4c
ms.date: 04/12/2019
localization_priority: Normal
---


# ShowNegativeBubbles property (Excel Graph)

**True** if negative bubbles are shown for the chart group. Valid only for bubble charts. Read/write **Boolean**.


## Syntax

_expression_.**ShowNegativeBubbles**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Example

This example makes negative bubbles visible for chart group one.

```vb
myChart.ChartGroups(1).ShowNegativeBubbles = True
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]