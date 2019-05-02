---
title: PivotCache.OptimizeCache property (Excel)
keywords: vbaxl10.chm227078
f1_keywords:
- vbaxl10.chm227078
ms.prod: excel
api_name:
- Excel.PivotCache.OptimizeCache
ms.assetid: 4aedf3bb-e15a-439c-5987-ea16cc233a7c
ms.date: 05/03/2019
localization_priority: Normal
---


# PivotCache.OptimizeCache property (Excel)

**True** if the PivotTable cache is optimized when it's constructed. The default value is **False**. Read/write **Boolean**.


## Syntax

_expression_.**OptimizeCache**

_expression_ A variable that represents a **[PivotCache](Excel.PivotCache.md)** object.


## Remarks

Cache optimization results in additional queries and degrades initial performance of the PivotTable report.

For OLE DB data sources, this property is read-only and always returns **False**.


## Example

This example causes the PivotTable cache for the first PivotTable report on worksheet one to be optimized when it's constructed.

```vb
Worksheets(1).PivotTables("Pivot1") _ 
 .PivotCache.OptimizeCache = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]