---
title: Shape.Shadow property (Word)
keywords: vbawd10.chm161480823
f1_keywords:
- vbawd10.chm161480823
ms.prod: word
api_name:
- Word.Shape.Shadow
ms.assetid: 43e65f16-9bd6-ab41-48b0-d52fc67dd5ae
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.Shadow property (Word)

Returns a  **ShadowFormat** object that represents the shadow formatting for the specified shape.


## Syntax

_expression_.**Shadow**

_expression_ Required. A variable that represents a **[Shape](Word.Shape.md)** object.


## Example

This example adds an arrow with shadow formatting to the active document.


```vb
Set myShape = ActiveDocument.Shapes _ 
 .AddShape(Type:=msoShapeRightArrow, _ 
 Left:=90, Top:=79, Width:=64, Height:=43) 
myShape.Shadow.Type = msoShadow5
```


## See also


[Shape Object](Word.Shape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]