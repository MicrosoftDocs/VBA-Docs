---
title: Scenarios.Item Method (Excel)
keywords: vbaxl10.chm362076
f1_keywords:
- vbaxl10.chm362076
ms.prod: excel
api_name:
- Excel.Scenarios.Item
ms.assetid: 6ed4b582-bd9c-5d18-f3ed-fc3b7b5a1580
ms.date: 06/08/2017
---


# Scenarios.Item Method (Excel)

Returns a single object from a collection.


## Syntax

 _expression_. `Item`( `_Index_` )

 _expression_ A variable that represents a [Scenarios](Excel.Scenarios.md) object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The name or index number for the object.|

### Return value

A  **[Scenario](Excel.Scenario.md)** object contained by the collection.


## Example

This example shows the scenario named Typical on the worksheet named Options.


```vb
Worksheets("options").Scenarios.Item("typical").Show
```


## See also


[Scenarios Object](Excel.Scenarios.md)

