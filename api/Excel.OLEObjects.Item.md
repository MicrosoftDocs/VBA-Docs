---
title: OLEObjects.Item method (Excel)
keywords: vbaxl10.chm422090
f1_keywords:
- vbaxl10.chm422090
ms.prod: excel
api_name:
- Excel.OLEObjects.Item
ms.assetid: 781b29f3-dcac-2679-72c2-a8d5d6280661
ms.date: 05/02/2019
localization_priority: Normal
---


# OLEObjects.Item method (Excel)

Returns a single object from a collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents an **[OLEObjects](Excel.OLEObjects.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The name or index number for the object.|

## Return value

An Object value that represents an object contained by the collection.


## Remarks

The text name of the object is the value of the **[Name](excel.name.name.md)** and **[Value](excel.name.value.md)** properties.


## Example

This example deletes OLE object one from Sheet1.

```vb
Worksheets("Sheet1").OLEObjects.Item(1).Delete
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]