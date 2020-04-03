---
title: Legend.IncludeInLayout property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.Legend.IncludeInLayout
ms.assetid: 2e14a6e0-923b-d383-2e40-dfa17f95df92
ms.date: 06/08/2017
localization_priority: Normal
---


# Legend.IncludeInLayout property (PowerPoint)

 **True** if a legend will occupy the chart layout space when a chart layout is being determined. The default is **True**. Read/write **Boolean**.


## Syntax

_expression_. `IncludeInLayout`

_expression_ A variable that represents a '[Legend](PowerPoint.Legend.md)' object.


## Remarks

This property does not affect whether a chart is in autolayout mode or not. If the user adds a title by using the  **Above Chart** command, the chart will resize smaller, as in previous versions of Microsoft Office. If the user then removes the title or selects one of the overlay title options, the chart will resize larger, as if the title were not on the chart.


## See also


[Legend Object](PowerPoint.Legend.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]