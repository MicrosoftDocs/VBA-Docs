---
title: ShapeRange.Item method (Excel)
keywords: vbaxl10.chm640074
f1_keywords:
- vbaxl10.chm640074
ms.prod: excel
api_name:
- Excel.ShapeRange.Item
ms.assetid: a8458e74-5279-3e47-308f-6c0647c00ee9
ms.date: 05/14/2019
localization_priority: Normal
---


# ShapeRange.Item method (Excel)

Returns a single object from a collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a **[ShapeRange](Excel.shaperange.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The name or index number for the object.|

## Return value

A **[Shape](Excel.Shape.md)** object contained by the collection.


## Example

This example sets the **[OnAction](excel.shape.onaction.md)** property for shape two in a shape range. If the `sr` variable doesn't represent a **ShapeRange** object, this example fails.

```vb
Dim sr As Shape 
sr.Item(2).OnAction = "ShapeAction"
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]