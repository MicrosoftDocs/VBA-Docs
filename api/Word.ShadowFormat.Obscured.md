---
title: ShadowFormat.Obscured property (Word)
keywords: vbawd10.chm164364389
f1_keywords:
- vbawd10.chm164364389
ms.prod: word
api_name:
- Word.ShadowFormat.Obscured
ms.assetid: 2746b925-a4f1-b5a6-04e5-7380ad79e20a
ms.date: 06/08/2017
localization_priority: Normal
---


# ShadowFormat.Obscured property (Word)

 **MsoTrue** if the shadow of the specified shape appears filled in and is obscured by the shape, even if the shape has no fill. **MsoFalse** if the shadow has no fill and the outline of the shadow is visible through the shape if the shape has no fill. Read/write **MsoTriState**.


## Syntax

_expression_.**Obscured**

_expression_ Required. A variable that represents a **[ShadowFormat](Word.ShadowFormat.md)** object.


## Example

This example sets the horizontal and vertical offsets for the shadow of shape three on myDocument. The shadow is offset 5 points to the right of the shape and 3 points above it. If the shape doesn't already have a shadow, this example adds one to it. The shadow will be filled in and obscured by the shape, even if the shape has no fill.


```vb
Set myDocument = ActiveDocument 
With myDocument.Shapes(3).Shadow 
 .Visible = True 
 .OffsetX = 5 
 .OffsetY = -3 
 .Obscured = msoTrue 
End With
```


## See also


[ShadowFormat Object](Word.ShadowFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]