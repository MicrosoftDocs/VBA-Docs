---
title: PivotTable.ClearAllFilters method (Excel)
keywords: vbaxl10.chm235170
f1_keywords:
- vbaxl10.chm235170
ms.prod: excel
api_name:
- Excel.PivotTable.ClearAllFilters
ms.assetid: e12fba36-f699-9800-99bc-d29b58b26043
ms.date: 05/08/2019
localization_priority: Normal
---


# PivotTable.ClearAllFilters method (Excel)

The **ClearAllFilters** method deletes all filters currently applied to the PivotTable. This includes deleting all filters in the **[PivotFilters](excel.pivotfilters.md)** collection, removing any manual filtering applied, and setting all PivotFields in the Report Filter area to the default item.


## Syntax

_expression_.**ClearAllFilters**

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Remarks

After calling the **ClearAllFilters** method, the **PivotFilters** collection will be empty.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
