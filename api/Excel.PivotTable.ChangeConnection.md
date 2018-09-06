---
title: PivotTable.ChangeConnection Method (Excel)
keywords: vbaxl10.chm235183
f1_keywords:
- vbaxl10.chm235183
ms.prod: excel
api_name:
- Excel.PivotTable.ChangeConnection
ms.assetid: 189c7ccc-d31c-dae8-f203-d590d1e46b82
ms.date: 06/08/2017
---


# PivotTable.ChangeConnection Method (Excel)

Changes the connection of the specified  **[PivotTable](Excel.PivotTable.md)** .


## Syntax

 _expression_. `ChangeConnection`( `_conn_` )

 _expression_ A variable that represents a [PivotTable](Excel.PivotTable.md) object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _conn_|Required| **WorkbookConnection**|A  **[WorkbookConnection](Excel.WorkbookConnection.md)** object that repesents the new conneciton for the PivotTable.|

## Remarks

The  **ChangeConnection** method can only be used with a **PivotTable** that is connected to an external data source. A run-time error will occur if the **ChangeConnection** method is used with a **PivotTable** that uses data stored on a worksheet as its data source.


## See also


[PivotTable Object](Excel.PivotTable.md)

