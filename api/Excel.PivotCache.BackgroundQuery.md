---
title: PivotCache.BackgroundQuery property (Excel)
keywords: vbaxl10.chm227073
f1_keywords:
- vbaxl10.chm227073
ms.prod: excel
api_name:
- Excel.PivotCache.BackgroundQuery
ms.assetid: 91909d27-68ca-a870-5cd9-72019c65f060
ms.date: 05/03/2019
localization_priority: Normal
---


# PivotCache.BackgroundQuery property (Excel)

**True** if queries for the PivotTable report are performed asynchronously (in the background). Read/write **Boolean**.


## Syntax

_expression_.**BackgroundQuery**

_expression_ A variable that represents a **[PivotCache](Excel.PivotCache.md)** object.


## Remarks

For OLAP data sources, this property is read-only and always returns **False**.


## Example

This example causes queries for the first PivotTable report on worksheet one to be performed in the background.

```vb
Worksheets(1).PivotTables("Pivot1") _ 
 .PivotCache.BackgroundQuery = True
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]