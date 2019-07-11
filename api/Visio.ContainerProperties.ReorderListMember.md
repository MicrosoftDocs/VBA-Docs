---
title: ContainerProperties.ReorderListMember method (Visio)
keywords: vis_sdr.chm17662340
f1_keywords:
- vis_sdr.chm17662340
ms.prod: visio
api_name:
- Visio.ContainerProperties.ReorderListMember
ms.assetid: 6bcb8928-750d-bea6-bee8-1a4f18cfd08e
ms.date: 06/08/2017
localization_priority: Normal
---


# ContainerProperties.ReorderListMember method (Visio)

Moves a shape or a set of shapes up or down in the list.


## Syntax

_expression_.**ReorderListMember** (_ObjectToReorder_, _Position_)

_expression_ A variable that represents a **[ContainerProperties](Visio.ContainerProperties.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ObjectToReorder_|Required| **[UNKNOWN]**|The shape or shapes to reorder in the container. Can be either **[Shape](Visio.Shape.md)** or **[Selection](Visio.Selection.md)** objects.|
| _Position_|Required| **Long**|The insertion position in the list, which is one-based.|

## Return value

**Nothing**


## Remarks

If the container is not a list, Microsoft Visio returns an Invalid Source error. 

If the  _ObjectToReorder_ parameter does not contain top-level shapes on the page, if any shape in _ObjectToReorder_ is not a member of the list, or if the list is locked, Visio returns an Invalid Parameter error.

To insert before the first item in the list, pass 1 for the  _Position_ parameter.

To insert after the final item in the list, set  _Position_ greater than or equal to the count of items.

If you pass an out-of-range value for  _Position_, Visio uses the nearest valid position.

If you pass a non-contiguous selection of list members for  _ObjectToReorder_, Visio makes the selection contiguous in the resulting reordered list, while maintaining relative position. For example, in a list ordered A,B,C,D, if you move B and D to position 1, the resulting list order is B,D,A,C.


## Example

The following Visual Basic for Applications (VBA) example shows how to use the **ReorderListMember** method to move a list member shape to the second position in the list.


```vb
vsoListShape.ContainerProperties.ReorderListMember vsoShape, 2
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]