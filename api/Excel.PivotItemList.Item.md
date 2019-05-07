---
title: PivotItemList.Item method (Excel)
keywords: vbaxl10.chm721074
f1_keywords:
- vbaxl10.chm721074
ms.prod: excel
api_name:
- Excel.PivotItemList.Item
ms.assetid: 69d0c71b-aa5a-b6cd-41d7-825197af869e
ms.date: 05/07/2019
localization_priority: Normal
---


# PivotItemList.Item method (Excel)

Returns a single object from a collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a **[PivotItemList](Excel.PivotItemList.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The name or index number for the object.|

## Return value

A **[PivotItem](Excel.PivotItem.md)** object contained by the collection.


## Remarks

The text name of the object is the value of the **[Name](Excel.PivotItem.Name.md)** and **[Value](Excel.PivotItem.Value.md)** properties.


## Example

This example hides calculated item one.

```vb
Worksheets(1).PivotTables(1).PivotFields("year") _ 
 .CalculatedItems.Item(1).Visible = False
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]