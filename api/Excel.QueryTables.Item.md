---
title: QueryTables.Item method (Excel)
keywords: vbaxl10.chm521075
f1_keywords:
- vbaxl10.chm521075
ms.prod: excel
api_name:
- Excel.QueryTables.Item
ms.assetid: c7b70ccd-1049-0d50-1536-f1d42b9b1e09
ms.date: 05/03/2019
localization_priority: Normal
---


# QueryTables.Item method (Excel)

Returns a single object from a collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a **[QueryTables](Excel.QueryTables.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The name or index number for the object.|

## Return value

A **[QueryTable](Excel.QueryTable.md)** object contained by the collection.


## Example

This example sets a query table so that formulas to the right of the query table are automatically updated whenever it's refreshed.

```vb
Sheets("sheet1").QueryTables.Item(1).FillAdjacentFormulas = True
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]