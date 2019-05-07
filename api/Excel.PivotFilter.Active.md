---
title: PivotFilter.Active property (Excel)
keywords: vbaxl10.chm770078
f1_keywords:
- vbaxl10.chm770078
ms.prod: excel
api_name:
- Excel.PivotFilter.Active
ms.assetid: 9fdbab3b-96e1-d821-5dc3-77a8a02c850a
ms.date: 05/07/2019
localization_priority: Normal
---


# PivotFilter.Active property (Excel)

Returns whether the specified PivotFilter is active. Read-only **Boolean**.


## Syntax

_expression_.**Active**

_expression_ A variable that represents a **[PivotFilter](Excel.PivotFilter.md)** object.


## Remarks

This property returns **True** when the PivotField of the filter is in the PivotTable and the filter is evaluated when the PivotTable is updated. It returns **False** when the PivotField of the filter is not in the PivotTable and has no effect on the PivotTable calculation.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]