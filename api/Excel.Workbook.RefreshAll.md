---
title: Workbook.RefreshAll method (Excel)
keywords: vbaxl10.chm199135
f1_keywords:
- vbaxl10.chm199135
api_name:
- Excel.Workbook.RefreshAll
ms.assetid: c1a956dc-263c-5c24-3b51-fc4af22dcd33
ms.date: 10/06/2026
ms.localizationpriority: medium
---


# Workbook.RefreshAll method (Excel)

Refreshes all external data ranges and PivotTable reports in the specified workbook.


## Syntax

_expression_.**RefreshAll**

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Remarks

Objects that have the **[BackgroundQuery](Excel.PivotCache.BackgroundQuery.md)** property set to **True** are refreshed in the background.

If a Power Query query that loads data into a worksheet table fails during refresh, **RefreshAll** doesn't raise a run-time error. Other connections are still refreshed, and the affected table retains its previous data. This behavior occurs even when background refresh is disabled for every connection.

To detect these errors, refresh each table's query separately by calling **[QueryTable.Refresh](Excel.QueryTable.Refresh.md)** with _BackgroundQuery_ set to **False**. Then, handle any error that the method raises. Use **ListObject.QueryTable** to get the query table for a table loaded by Power Query.


## Example

This example refreshes all external data ranges and PivotTable reports in the third workbook.

```vb
Workbooks(3).RefreshAll
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
