---
title: ChartTitle.IncludeInLayout property (Excel)
keywords: vbaxl10.chm563092
f1_keywords:
- vbaxl10.chm563092
api_name:
- Excel.ChartTitle.IncludeInLayout
ms.assetid: 29a38d5a-9aaa-bcbc-7a86-96ce85286cf1
ms.date: 04/20/2019
ms.localizationpriority: medium
---


# ChartTitle.IncludeInLayout property (Excel)

**True** if a chart title will occupy the chart layout space when a chart layout is being determined. The default value is **True**. Read/write **Boolean**.


## Syntax

_expression_.**IncludeInLayout**

_expression_ A variable that represents a **[ChartTitle](Excel.ChartTitle(object).md)** object.


## Remarks

This property does not affect whether a chart is in autolayout mode or not. If the user adds a title by using the **Above Chart** command, the chart will resize smaller. If the user then removes the title or selects one of the overlay title options, the chart will resize larger, as if the title were not on the chart.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]