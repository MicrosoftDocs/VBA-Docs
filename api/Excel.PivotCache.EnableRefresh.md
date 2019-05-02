---
title: PivotCache.EnableRefresh property (Excel)
keywords: vbaxl10.chm227075
f1_keywords:
- vbaxl10.chm227075
ms.prod: excel
api_name:
- Excel.PivotCache.EnableRefresh
ms.assetid: 5919198f-bb4a-eb54-1a28-41033b525fa1
ms.date: 05/03/2019
localization_priority: Normal
---


# PivotCache.EnableRefresh property (Excel)

**True** if the PivotTable cache or query table can be refreshed by the user. The default value is **True**. Read/write **Boolean**.


## Syntax

_expression_.**EnableRefresh**

_expression_ A variable that represents a **[PivotCache](Excel.PivotCache.md)** object.


## Remarks

The **[RefreshOnFileOpen](Excel.PivotCache.RefreshOnFileOpen.md)** property is ignored if the **EnableRefresh** property is set to **False**.

For OLAP data sources, setting this property to **False** disables updates.


## Example

This example sets the first PivotTable report on worksheet one so that it cannot be refreshed.

```vb
Worksheets(1).PivotTables("Pivot1") _ 
 .PivotCache.EnableRefresh = False
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]