---
title: ServerViewableItems.Add method (Excel)
keywords: vbaxl10.chm833074
f1_keywords:
- vbaxl10.chm833074
ms.prod: excel
api_name:
- Excel.ServerViewableItems.Add
ms.assetid: e5771bed-efd0-3cdc-ce80-13b71f596d01
ms.date: 06/08/2017
localization_priority: Normal
---


# ServerViewableItems.Add method (Excel)

Adds a reference to the  **[ServerViewableItems](Excel.ServerViewableItems.md)** collection.


## Syntax

_expression_. `Add`( `_Obj_` )

_expression_ A variable that represents a [ServerViewableItems](./Excel.ServerViewableItems.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Obj_|Required| **Variant**|The reference to an object. The object can be a reference to sheets or named items (for example, named ranges, charts, tables, and PivotTables). You cannot have both sheets and named items in the same collection.|

## Return value

Object


## Remarks

If you try to add a mix of both sheets and named items to the  **[ServerViewableItems](Excel.ServerViewableItems.md)** collection, an error is returned. The **[ServerViewableItems](Excel.ServerViewableItems.md)** collection can contain references only to sheets, or references only to named items, but not both in the same call.


## See also


[ServerViewableItems Object](Excel.ServerViewableItems.md)

