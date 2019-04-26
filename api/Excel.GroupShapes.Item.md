---
title: GroupShapes.Item method (Excel)
keywords: vbaxl10.chm642074
f1_keywords:
- vbaxl10.chm642074
ms.prod: excel
api_name:
- Excel.GroupShapes.Item
ms.assetid: 2f664b81-d870-7936-6dff-0eef163c2e03
ms.date: 04/26/2019
localization_priority: Normal
---


# GroupShapes.Item method (Excel)

Returns a single object from a collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a **[GroupShapes](Excel.GroupShapes.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The name or index number for the object.|

## Return value

A **[Shape](Excel.Shape.md)** object contained by the collection.


## Example

This example sets the **OnAction** property for shape two in a shape range. If the `sr` variable doesn't represent a **ShapeRange** object, this example fails.

```vb
Dim sr As Shape 
sr.Item(2).OnAction = "ShapeAction"
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]