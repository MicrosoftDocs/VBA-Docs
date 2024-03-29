---
title: ChartGroup.VaryByCategories property (Excel)
keywords: vbaxl10.chm568093
f1_keywords:
- vbaxl10.chm568093
api_name:
- Excel.ChartGroup.VaryByCategories
ms.assetid: 9ae94a48-abc7-b692-7376-f4cb780a4063
ms.date: 04/20/2019
ms.localizationpriority: medium
---


# ChartGroup.VaryByCategories property (Excel)

**True** if Microsoft Excel assigns a different color or pattern to each data marker. The chart must contain only one series. Read/write **Boolean**.


## Syntax

_expression_.**VaryByCategories**

_expression_ A variable that represents a **[ChartGroup](Excel.ChartGroup(object).md)** object.


## Example

This example assigns a different color or pattern to each data marker in chart group one. The example should be run on a 2D line chart that has data markers on a series.

```vb
Charts("Chart1").ChartGroups(1).VaryByCategories = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]