---
title: PivotCaches.Item method (Excel)
keywords: vbaxl10.chm229074
f1_keywords:
- vbaxl10.chm229074
ms.prod: excel
api_name:
- Excel.PivotCaches.Item
ms.assetid: 80a830fb-a1bf-f1dd-962c-339d99b6f80d
ms.date: 05/03/2019
localization_priority: Normal
---


# PivotCaches.Item method (Excel)

Returns a single object from a collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a **[PivotCaches](Excel.PivotCaches.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The name or index number for the object.|

## Return value

A **[PivotCache](Excel.PivotCache.md)** object contained by the collection.


## Example

This example refreshes cache one.

```vb
ActiveWorkbook.PivotCaches.Item(1).Refresh
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]