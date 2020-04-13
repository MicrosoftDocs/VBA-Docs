---
title: SeriesCollection.Creator property (Project)
ms.prod: project-server
ms.assetid: d2bc1554-6ae3-7eb2-e455-fef0cf544290
ms.date: 06/08/2017
localization_priority: Normal
---


# SeriesCollection.Creator property (Project)
Gets a 32-bit integer that indicates the application in which the series collection was created. Read-only  **Long**.

## Syntax

_expression_.**Creator**

_expression_ A variable that represents a 'SeriesCollection' object.


## Remarks

If the chart was created in Microsoft Project, the **Creator** property returns the decimal number `1347571530`, which is hexadecimal  `0x50524F4A`, which is equivalent to the string  **PROJ**. For example, run the following command in the Immediate window of the VBE, with the name of the active report.


```vb
? ActiveProject.Reports("Simple scalar chart").Shapes(1).Chart.SeriesCollection.Creator
```


## See also


[SeriesCollection Object](Project.seriescollection.md)
[Chart.Creator Property](Project.chart.creator.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]