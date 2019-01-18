---
title: Shape.Shadow Property (Publisher)
keywords: vbapb10.chm2228296
f1_keywords:
- vbapb10.chm2228296
ms.prod: publisher
api_name:
- Publisher.Shape.Shadow
ms.assetid: cfb908ae-ef1d-9539-1f82-2693cbe38d97
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.Shadow Property (Publisher)

Returns a  **[ShadowFormat](Publisher.ShadowFormat.md)** object that represents the shadow formatting for the specified shape.


## Syntax

 _expression_. **Shadow**

 _expression_ A variable that represents a  **Shape** object.


## Example

This example adds an arrow with shadow formatting and fill color to the first page in the active document.


```vb
Sub SetShapeShadow() 
 With ActiveDocument.Pages(1).Shapes.AddShape( _ 
 Type:=msoShapeRightArrow, Left:=72, _ 
 Top:=72, Width:=64, Height:=43) 
 .Shadow.Type = msoShadow5 
 .Fill.ForeColor.RGB = RGB(Red:=255, Green:=0, Blue:=255) 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]