---
title: PivotTable.CommitChanges method (Excel)
keywords: vbaxl10.chm235192
f1_keywords:
- vbaxl10.chm235192
api_name:
- Excel.PivotTable.CommitChanges
ms.assetid: f64031c6-8309-7c8a-5786-949d2ec10dea
ms.date: 05/08/2019
ms.localizationpriority: medium
---


# PivotTable.CommitChanges method (Excel)

Performs a commit operation on the data source of a PivotTable report based on an OLAP data source.


## Syntax

_expression_.**CommitChanges**

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Return value

Nothing


## Remarks

The **CommitChanges** method sends a **COMMIT TRANSACTION** statement to the OLAP server, and clears all cells that were edited by entering a value, but will not clear formulas in value cells. This method generates a run-time error if it is executed on a PivotTable report based on a non-OLAP data source.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]