---
title: AxisTitle.IncludeInLayout property (Excel)
keywords: vbaxl10.chm566075
f1_keywords:
- vbaxl10.chm566075
ms.prod: excel
api_name:
- Excel.AxisTitle.IncludeInLayout
ms.assetid: ef84d235-6d60-f5c9-f185-e474a8b6a0e7
ms.date: 04/13/2019
localization_priority: Normal
---


# AxisTitle.IncludeInLayout property (Excel)

**True** if an axis title will occupy the chart layout space when a chart layout is being determined. The default value is **True**. Read/write **Boolean**.


## Syntax

_expression_.**IncludeInLayout**

_expression_ A variable that represents an **[AxisTitle](Excel.AxisTitle(object).md)** object.


## Remarks

This property does not affect whether a chart is in autolayout mode. If the user adds a title by using the **Above Chart** command, the chart will resize smaller. If the user then removes the title or selects one of the overlay title options, the chart will resize larger, as if the title were not on the chart.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]