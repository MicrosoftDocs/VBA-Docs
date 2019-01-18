---
title: Workbook.RefreshAll method (Excel)
keywords: vbaxl10.chm199135
f1_keywords:
- vbaxl10.chm199135
ms.prod: excel
api_name:
- Excel.Workbook.RefreshAll
ms.assetid: c1a956dc-263c-5c24-3b51-fc4af22dcd33
ms.date: 06/08/2017
localization_priority: Priority
---


# Workbook.RefreshAll method (Excel)

Refreshes all external data ranges and PivotTable reports in the specified workbook.


## Syntax

_expression_. `RefreshAll`

_expression_ A variable that represents a [Workbook](./Excel.Workbook.md) object.


## Remarks

Objects that have the  **[BackgroundQuery](Excel.PivotCache.BackgroundQuery.md)** property set to **True** are refreshed in the background.


## Example

This example refreshes all external data ranges and PivotTable reports in the third workbook.


```vb
Workbooks(3).RefreshAll
```


## See also


[Workbook Object](Excel.Workbook.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]