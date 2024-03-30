---
title: ChartGroup.SizeRepresents property (Excel)
keywords: vbaxl10.chm568094
f1_keywords:
- vbaxl10.chm568094
api_name:
- Excel.ChartGroup.SizeRepresents
ms.assetid: db7811b5-6d65-d3e0-0c0b-83dcd3692cf1
ms.date: 04/20/2019
ms.localizationpriority: medium
---


# ChartGroup.SizeRepresents property (Excel)

Returns or sets what the bubble size represents on a bubble chart. Can be either of the following **[XlSizeRepresents](Excel.XlSizeRepresents.md)** constants: **xlSizeIsArea** or **xlSizeIsWidth**. Read/write **Long**.


## Syntax

_expression_.**SizeRepresents**

_expression_ A variable that represents a **[ChartGroup](Excel.ChartGroup(object).md)** object.


## Example

This example sets what the bubble size represents for chart group one.

```vb
Charts(1).ChartGroups(1).SizeRepresents = xlSizeIsWidth
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]