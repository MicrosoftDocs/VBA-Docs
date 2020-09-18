---
title: PivotTable.ChangePivotCache method (Excel)
keywords: vbaxl10.chm235184
f1_keywords:
- vbaxl10.chm235184
ms.prod: excel
api_name:
- Excel.PivotTable.ChangePivotCache
ms.assetid: 1b1ee1b4-0ed6-641a-3e1d-739461fa0466
ms.date: 05/08/2019
localization_priority: Normal
---


# PivotTable.ChangePivotCache method (Excel)

Changes the **[PivotCache](Excel.PivotCache.md)** object of the specified **PivotTable**.


## Syntax

_expression_.**ChangePivotCache** (_bstr_)

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _bstr_|Required| **String**|A **PivotTable** or **PivotCache** object that represents the new PivotCache for the specified PivotTable.|

## Remarks

The **ChangePivotCache** method can only be used with a PivotTable that uses data stored on a worksheet as its data source. A run-time error occurs if the **ChangePivotCache** method is used with a PivotTable that is connected to an external data source.


## Example

In the following code sample, the pivot table named **PivotTable1** is on Sheet1.  The code changes its pivot cache to a cache created from the data stored in the table called **Table2** in the same workbook.

```vb
Sheets("Sheet1").PivotTables("PivotTable1").ChangePivotCache _
   ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:="Table2", Version:=xlPivotTableVersion15)
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
