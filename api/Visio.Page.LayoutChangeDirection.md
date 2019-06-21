---
title: Page.LayoutChangeDirection method (Visio)
keywords: vis_sdr.chm10962145
f1_keywords:
- vis_sdr.chm10962145
ms.prod: visio
api_name:
- Visio.Page.LayoutChangeDirection
ms.assetid: f818785b-d845-34de-50d1-e68c3c09dda9
ms.date: 06/08/2017
localization_priority: Normal
---


# Page.LayoutChangeDirection method (Visio)

Revises the layout of a set of connected shapes on the page, by rotating or flipping a connected diagram without rotating or flipping the individual shapes.


## Syntax

_expression_. `LayoutChangeDirection`( `_Direction_` )

_expression_ A variable that represents a **[Page](Visio.Page.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Direction_|Required| **[VisLayoutDirection](Visio.VisLayoutDirection.md)**|The action to take. See Remarks for possible values.|

## Return value

 **Nothing**


## Remarks

The  _Direction_ parameter must be one of the following **VisLayoutDirection** constants.



|Constant|Value|Description|
|:-----|:-----|:-----|
| **visLayoutDirRotateRight**|0|Rotates the diagram 90 degrees clockwise.|
| **visLayoutDirRotateLeft**|1|Rotates the diagram 90 degrees counterclockwise.|
| **visLayoutDirFlipVert**|2|Flips the diagram vertically.|
| **visLayoutDirFlipHorz**|3|Flips the diagram horizontally.|

## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **LayoutChangeDirection** method to flip connected shapes on the active page vertically, without flipping the individual shapes.


```vb
Public Sub PageLayoutChangeDirection_Example()
   ActivePage.LayoutChangeDirection (visLayoutDirFlipVert)
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]