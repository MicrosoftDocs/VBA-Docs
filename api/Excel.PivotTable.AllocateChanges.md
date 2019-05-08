---
title: PivotTable.AllocateChanges method (Excel)
keywords: vbaxl10.chm235191
f1_keywords:
- vbaxl10.chm235191
ms.prod: excel
api_name:
- Excel.PivotTable.AllocateChanges
ms.assetid: 6eb2d6b6-7340-fe63-611c-0972b9ccf496
ms.date: 05/08/2019
localization_priority: Normal
---


# PivotTable.AllocateChanges method (Excel)

Performs a writeback operation for all edited cells in a PivotTable report based on an OLAP data source.


## Syntax

_expression_.**AllocateChanges**

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Return value

Nothing


## Remarks

The **AllocateChanges** method will execute an **UPDATE CUBE** statement for all changes made in the values area of the PivotTable since the last apply changes operation was committed, or since the PivotTable was created if commiting apply changes has never been performed. 

This method generates a run-time error if it is executed on a PivotTable report based on a non-OLAP data source.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]