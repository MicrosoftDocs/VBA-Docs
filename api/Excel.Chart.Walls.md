---
title: Chart.Walls property (Excel)
keywords: vbaxl10.chm149151
f1_keywords:
- vbaxl10.chm149151
ms.prod: excel
api_name:
- Excel.Chart.Walls
ms.assetid: fbee1165-7602-4d77-e5b6-8a127783c96e
ms.date: 04/16/2019
localization_priority: Normal
---


# Chart.Walls property (Excel)

Returns a **[Walls](Excel.Walls(object).md)** object that represents the walls of the 3D chart. Read-only.


## Syntax

_expression_.**Walls**

_expression_ A variable that represents a **[Chart](Excel.Chart(object).md)** object.


## Example

This example sets the color of the wall border of Chart1 to red. The example should be run on a 3D chart.

```vb
Charts("Chart1").Walls.Border.ColorIndex = 3
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]