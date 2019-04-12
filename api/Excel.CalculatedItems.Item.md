---
title: CalculatedItems.Item method (Excel)
keywords: vbaxl10.chm250075
f1_keywords:
- vbaxl10.chm250075
ms.prod: excel
api_name:
- Excel.CalculatedItems.Item
ms.assetid: ad7642b5-2579-17b4-ed2f-ebcac54bb595
ms.date: 04/13/2019
localization_priority: Normal
---


# CalculatedItems.Item method (Excel)

Returns a single object from a collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a **[CalculatedItems](Excel.CalculatedItems.md)** object.


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