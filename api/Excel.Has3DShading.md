---
title: Has3DShading property (Excel Graph)
keywords: vbagr10.chm5207443
f1_keywords:
- vbagr10.chm5207443
ms.prod: excel
api_name:
- Excel.Has3DShading
ms.assetid: 1a6d41c5-83d5-72f6-f8d5-86cbf52af501
ms.date: 04/11/2019
localization_priority: Normal
---


# Has3DShading property (Excel Graph)

**True** if the chart group has three-dimensional shading. Read/write **Boolean**.

## Syntax

_expression_.**Has3DShading**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Example

This example adds three-dimensional shading to chart group one on the chart.

```vb
Charts(1).ChartGroups(1).Has3DShading = True
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]