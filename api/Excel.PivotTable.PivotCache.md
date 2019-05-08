---
title: PivotTable.PivotCache method (Excel)
keywords: vbaxl10.chm235115
f1_keywords:
- vbaxl10.chm235115
ms.prod: excel
api_name:
- Excel.PivotTable.PivotCache
ms.assetid: 82602154-783d-3f78-b354-0dabfdc34c98
ms.date: 05/09/2019
localization_priority: Normal
---


# PivotTable.PivotCache method (Excel)

Returns a **[PivotCache](Excel.PivotCache.md)** object that represents the cache for the specified PivotTable report. Read-only.


## Syntax

_expression_.**PivotCache**

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Return value

PivotCache


## Example

This example causes the PivotTable cache for the first PivotTable report on worksheet one to be optimized when it's constructed.

```vb
Worksheets(1).PivotTables("Pivot1") _ 
 .PivotCache.OptimizeCache = True 

```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]