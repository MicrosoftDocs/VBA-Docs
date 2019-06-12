---
title: ShadowFormat object (Publisher)
keywords: vbapb10.chm3735551
f1_keywords:
- vbapb10.chm3735551
ms.prod: publisher
api_name:
- Publisher.ShadowFormat
ms.assetid: b23ab92e-5e49-8d8d-69d5-93d391a9edb2
ms.date: 06/01/2019
localization_priority: Normal
---


# ShadowFormat object (Publisher)

Represents shadow formatting for a shape.
 
## Remarks

Use the **[Shadow](publisher.shape.shadow.md)** property of the **Shape** or **[ShapeRange](publisher.shaperange.shadow.md)** object to return a **ShadowFormat** object. 

## Example

The following example adds a shadowed rectangle to the active document. The pink shadow is offset 7 points to the right of the rectangle and 7 points above it.

```vb
Sub FormatShadow() 
 With ActiveDocument.Pages(1).Shapes.AddShape( _ 
 Type:=msoShapeRectangle, Left:=72, Top:=72, _ 
 Width:=100, Height:=200).Shadow 
 .ForeColor.RGB = RGB(Red:=255, Green:=0, Blue:=150) 
 .Obscured = msoTrue 
 .OffsetX = 7 
 .OffsetY = -7 
 .Visible = True 
 End With 
End Sub
```


## Methods

- [IncrementOffsetX](Publisher.ShadowFormat.IncrementOffsetX.md)
- [IncrementOffsetY](Publisher.ShadowFormat.IncrementOffsetY.md)

## Properties

- [Application](Publisher.ShadowFormat.Application.md)
- [Blur](Publisher.ShadowFormat.Blur.md)
- [ForeColor](Publisher.ShadowFormat.ForeColor.md)
- [Obscured](Publisher.ShadowFormat.Obscured.md)
- [OffsetX](Publisher.ShadowFormat.OffsetX.md)
- [OffsetY](Publisher.ShadowFormat.OffsetY.md)
- [Parent](Publisher.ShadowFormat.Parent.md)
- [RotateWithShape](Publisher.shadowformat.rotatewithshape.md)
- [Size](Publisher.ShadowFormat.Size.md)
- [Type](Publisher.ShadowFormat.Type.md)
- [Visible](Publisher.ShadowFormat.Visible.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]