---
title: Axis.TickLabels property (Excel)
keywords: vbaxl10.chm561100
f1_keywords:
- vbaxl10.chm561100
api_name:
- Excel.Axis.TickLabels
ms.assetid: e8a6edf9-2fdd-d8e9-0de9-5c4aa921c6b1
ms.date: 04/13/2019
ms.localizationpriority: medium
---


# Axis.TickLabels property (Excel)

Returns a **[TickLabels](Excel.TickLabels(object).md)** object that represents the tick-mark labels for the specified axis. Read-only.


## Syntax

_expression_.**TickLabels**

_expression_ A variable that represents an **[Axis](Excel.Axis(object).md)** object.


## Example

This example sets the color of the tick-mark label font for the value axis on Chart1.

```vb
Charts("Chart1").Axes(xlValue).TickLabels.Font.ColorIndex = 3
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]