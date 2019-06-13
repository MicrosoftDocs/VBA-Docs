---
title: ShadowFormat.OffsetX property (Publisher)
keywords: vbapb10.chm3670274
f1_keywords:
- vbapb10.chm3670274
ms.prod: publisher
api_name:
- Publisher.ShadowFormat.OffsetX
ms.assetid: 2b34ace8-5c3b-002b-df96-43c8aef2fbd2
ms.date: 06/13/2019
localization_priority: Normal
---


# ShadowFormat.OffsetX property (Publisher)

Returns or sets a **Variant** value indicating the horizontal offset of the shadow from the specified shape. A positive value offsets the shadow to the right of the shape; a negative value offsets it to the left. Read/write.


## Syntax

_expression_.**OffsetX**

_expression_ A variable that represents a **[ShadowFormat](Publisher.ShadowFormat.md)** object.


## Return value

Variant


## Remarks

Numeric values are evaluated in [points](../language/glossary/vbe-glossary.md#point); strings can be in any units supported by Microsoft Publisher (for example, "2.5 in").

If you want to nudge a shadow horizontally or vertically from its current position without having to specify an absolute position, use the **[IncrementOffsetX](Publisher.ShadowFormat.IncrementOffsetX.md)** method or the **[IncrementOffsetY](Publisher.ShadowFormat.IncrementOffsetY.md)** method.


## Example

This example sets the horizontal and vertical offsets of the shadow for shape three on page one of the active publication. The shadow is offset 5 points to the right of the shape and 3 points above it. If the shape doesn't already have a shadow, this example adds one to it.

```vb
With ActiveDocument.Pages(1).Shapes(3).Shadow 
 .Visible = True 
 .OffsetX = 5 
 .OffsetY = -3 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]