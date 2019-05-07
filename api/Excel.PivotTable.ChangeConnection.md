---
title: PivotTable.ChangeConnection method (Excel)
keywords: vbaxl10.chm235183
f1_keywords:
- vbaxl10.chm235183
ms.prod: excel
api_name:
- Excel.PivotTable.ChangeConnection
ms.assetid: 189c7ccc-d31c-dae8-f203-d590d1e46b82
ms.date: 05/08/2019
localization_priority: Normal
---


# PivotTable.ChangeConnection method (Excel)

Changes the connection of the specified **PivotTable** object.


## Syntax

_expression_.**ChangeConnection** (_conn_)

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _conn_|Required| **[WorkbookConnection](Excel.WorkbookConnection.md)** |A **WorkbookConnection** object that represents the new connection for the PivotTable.|

## Remarks

The **ChangeConnection** method can only be used with a PivotTable that is connected to an external data source. A run-time error occurs if the **ChangeConnection** method is used with a PivotTable that uses data stored on a worksheet as its data source.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]