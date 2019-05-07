---
title: PivotTable.CalculatedMembersInFilters property (Excel)
keywords: vbaxl10.chm235202
f1_keywords:
- vbaxl10.chm235202
ms.prod: excel
api_name:
- Excel.PivotTable.CalculatedMembersInFilters
ms.assetid: 1f28b21d-d079-e37a-563e-473e6b57bccd
ms.date: 05/08/2019
localization_priority: Normal
---


# PivotTable.CalculatedMembersInFilters property (Excel)

Returns or sets whether to evaluate calculated members from OLAP servers in filters. Read/write.


## Syntax

_expression_.**CalculatedMembersInFilters**

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Return value

**Boolean**


## Remarks

**True** if calculated members are evaluated in filters; otherwise, **False**.

The value of this property corresponds to the setting of the **Evaluate calculated members from OLAP servers in filters** check box on the **Totals & Filters** tab of the **PivotTable Options** dialog box for a PivotTable report based on an OLAP data source.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]