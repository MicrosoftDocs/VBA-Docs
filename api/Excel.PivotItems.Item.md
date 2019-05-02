---
title: PivotItems.Item method (Excel)
keywords: vbaxl10.chm248076
f1_keywords:
- vbaxl10.chm248076
ms.prod: excel
api_name:
- Excel.PivotItems.Item
ms.assetid: 2ce0e125-1613-4dd9-9afa-623f6cca52b7
ms.date: 06/08/2017
localization_priority: Normal
---


# PivotItems.Item method (Excel)

Returns a single object from a collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a **[PivotItems](Excel.PivotItems.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The name or index number for the object.|

## Return value

An Object value that represents an object contained by the collection.


## Remarks

The text name of the object is the value of the  **Name** and **Value** properties.


## Example

This example hides the 1998 item in the first PivotTable report on Sheet3.


```vb
Worksheets("sheet3").PivotTables(1) _ 
 .PivotFields("year").PivotItems.Item("1998").Visible = False
```


## See also


[PivotItems Object](Excel.PivotItems.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]