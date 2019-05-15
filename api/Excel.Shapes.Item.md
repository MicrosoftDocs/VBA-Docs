---
title: Shapes.Item method (Excel)
keywords: vbaxl10.chm638074
f1_keywords:
- vbaxl10.chm638074
ms.prod: excel
api_name:
- Excel.Shapes.Item
ms.assetid: efd7e247-5976-95b1-3365-34997feb323f
ms.date: 05/15/2019
localization_priority: Normal
---


# Shapes.Item method (Excel)

Returns a single object from a collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a **[Shapes](Excel.Shapes.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The name or index number for the object.|

## Return value

A **[Shape](Excel.Shape.md)** object contained by the collection.


## Remarks

The text name of the object is the value of the **[Name](Excel.Shape.Name.md)** property.


## Example

This example sets the **[OnAction](excel.shape.onaction.md)** property for shape two in a **Shapes** collection. If the `ss` variable doesn't represent a **Shapes** object, this example fails.

```vb
Dim ss As Shape 
ss.Item(2).OnAction = "ShapeAction"
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
