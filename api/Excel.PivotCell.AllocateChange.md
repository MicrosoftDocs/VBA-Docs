---
title: PivotCell.AllocateChange method (Excel)
keywords: vbaxl10.chm692085
f1_keywords:
- vbaxl10.chm692085
ms.prod: excel
api_name:
- Excel.PivotCell.AllocateChange
ms.assetid: 21865f48-a011-478b-b485-16eba786dd92
ms.date: 05/04/2019
localization_priority: Normal
---


# PivotCell.AllocateChange method (Excel)

Performs a writeback operation on the specified cell in a PivotTable report based on an OLAP data source.


## Syntax

_expression_.**AllocateChange**

_expression_ A variable that represents a **[PivotCell](Excel.PivotCell.md)** object.


## Return value

**Nothing**


## Remarks

This method executes an **UPDATE CUBE** statement to add just the change in this particular cell, but also includes any previous changes applied. After the **UPDATE CUBE** statement is executed, a PivotTable update is run, and then a **ROLLBACK TRANSACTION** statement is executed.

Running the **AllocateChange** method for a cell in a PivotTable report based on a non-OLAP data source generates a run-time error.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]