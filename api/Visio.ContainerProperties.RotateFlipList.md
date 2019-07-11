---
title: ContainerProperties.RotateFlipList method (Visio)
keywords: vis_sdr.chm17662360
f1_keywords:
- vis_sdr.chm17662360
ms.prod: visio
api_name:
- Visio.ContainerProperties.RotateFlipList
ms.assetid: 0402f4e3-e494-b915-e6c3-a09a7fc12845
ms.date: 06/08/2017
localization_priority: Normal
---


# ContainerProperties.RotateFlipList method (Visio)

Rotates or flips the list direction for a list of shapes.


## Syntax

_expression_.**RotateFlipList** (_Direction_)

_expression_ A variable that represents a **[ContainerProperties](Visio.ContainerProperties.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Direction_|Required| **[VisLayoutDirection](Visio.VisLayoutDirection.md)**|The layout action to take.|

## Return value

**Nothing**


## Remarks

If the list contains container shapes only, and no other shapes, and if the ObjType ShapeSheet cell value of the list shape equals zero (0), nothing happens.

If the list contains container shapes only, and no other shapes, and if the ObjType ShapeSheet cell value of the list shape does not equal zero (0), the **RotateFlipList** method also rotates or flips the contents of the container shapes.

If the list contains a mix of container and non-container shapes, the method does not rotate or flip the contents of the containers but, rather, rotates or flips the entire list.

You can also use this method on lists that are paired in an overlapped list relationship. For rotation, both overlapped lists are rotated by 90 degrees. For flip, the overlapped list direction is not changed.

If the container is not a list, Microsoft Visio returns an Invalid Source error.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]