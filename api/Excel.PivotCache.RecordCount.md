---
title: PivotCache.RecordCount property (Excel)
keywords: vbaxl10.chm227079
f1_keywords:
- vbaxl10.chm227079
ms.prod: excel
api_name:
- Excel.PivotCache.RecordCount
ms.assetid: 5fcdcf2d-d52f-6ac1-ef09-8377fc5a1f4d
ms.date: 05/03/2019
localization_priority: Normal
---


# PivotCache.RecordCount property (Excel)

Returns the number of records in the PivotTable cache or the number of cache records that contain the specified item. Read-only **Long**.


## Syntax

_expression_.**RecordCount**

_expression_ A variable that represents a **[PivotCache](Excel.PivotCache.md)** object.


## Remarks

This property reflects the transient state of the cache at the time that it's queried. The cache can change between queries.


## Example

This example displays the number of cache records that contain Kiwi in the Products field.

```vb
MsgBox Worksheets(1).PivotTables("Pivot1") _ 
 .PivotFields("Product").PivotItems("Kiwi").RecordCount
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]