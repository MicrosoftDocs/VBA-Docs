---
title: CalculatedFields.Item method (Excel)
keywords: vbaxl10.chm244075
f1_keywords:
- vbaxl10.chm244075
ms.prod: excel
api_name:
- Excel.CalculatedFields.Item
ms.assetid: cae0c3a5-3403-f1b1-fe7f-c38ff6be6b07
ms.date: 04/13/2019
localization_priority: Normal
---


# CalculatedFields.Item method (Excel)

Returns a single object from a collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a **[CalculatedFields](Excel.CalculatedFields.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The name or index number for the object.|

## Return value

A **[PivotField](Excel.PivotField.md)** object contained by the collection.


## Remarks

The text name of the object is the value of the **[Name](Excel.PivotField.Name.md)** and **[Value](Excel.PivotField.Value.md)** properties.


## Example

This example sets the formula for calculated field one.

```vb
Worksheets(1).PivotTables(1).CalculatedFields.Item(1) _ 
 .Formula = "=Revenue - Cost"
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]