---
title: ShapeRange.GroupItems property (Word)
keywords: vbawd10.chm162857068
f1_keywords:
- vbawd10.chm162857068
ms.prod: word
api_name:
- Word.ShapeRange.GroupItems
ms.assetid: 800c95fd-2306-f614-d8b5-6a087eb3a2dc
ms.date: 06/08/2017
localization_priority: Normal
---


# ShapeRange.GroupItems property (Word)

Returns a  **[GroupShapes](Word.groupshapes.md)** object that represents the individual shapes in the specified group. Read-only.


## Syntax

_expression_. `GroupItems`

_expression_ A variable that represents a **[ShapeRange](Word.shaperange.md)** object.


## Remarks

 This property applies to **ShapeRange** objects that represent grouped shapes. Use the **Item** method of the **[GroupShapes](Word.groupshapes.md)** object to return a single shape from the group.


## See also


[ShapeRange Collection Object](Word.shaperange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]