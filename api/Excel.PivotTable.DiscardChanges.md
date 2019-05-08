---
title: PivotTable.DiscardChanges method (Excel)
keywords: vbaxl10.chm235193
f1_keywords:
- vbaxl10.chm235193
ms.prod: excel
api_name:
- Excel.PivotTable.DiscardChanges
ms.assetid: 9ee2905f-7dd1-81d2-7075-7fdc78ad6f1c
ms.date: 05/08/2019
localization_priority: Normal
---


# PivotTable.DiscardChanges method (Excel)

Discards all changes in the edited cells of a PivotTable report based on an OLAP data source.


## Syntax

_expression_.**DiscardChanges**

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Return value

Nothing


## Remarks

For a PivotTable report based on an OLAP data source, the method removes all values and formulas entered in value cells, and then runs a PivotTable update operation to retrieve the latest values from the data source. It sets the data source value to **NULL** for all value cells that are edited, and also executes a **ROLLBACK TRANSACTION** statement against the OLAP server.

If you try to execute this method for a PivotTable report based on a non-OLAP data source, this method generates a run-time error.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]