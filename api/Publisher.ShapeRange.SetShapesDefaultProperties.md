---
title: ShapeRange.SetShapesDefaultProperties method (Publisher)
keywords: vbapb10.chm2293800
f1_keywords:
- vbapb10.chm2293800
ms.prod: publisher
api_name:
- Publisher.ShapeRange.SetShapesDefaultProperties
ms.assetid: 1146cbf8-6d31-9fb8-c6a4-d54b68436cbd
ms.date: 06/14/2019
localization_priority: Normal
---


# ShapeRange.SetShapesDefaultProperties method (Publisher)

Applies the formatting for the specified shape or shape range to the default shape. Shapes created after this method has been used have this formatting applied to them by default.


## Syntax

_expression_.**SetShapesDefaultProperties**

_expression_ A variable that represents a **[ShapeRange](Publisher.ShapeRange.md)** object.


## Return value

Nothing


## Remarks

The **SetShapesDefaultProperties** method stores two different sets of default properties, one for a **Shape** object's **[AutoShapeType](Publisher.Shape.AutoShapeType.md)** property, and another for a **[TextFrame](Publisher.TextFrame.md)** object. 

In other words, if this method is called on an AutoShape, the default formatting of that object will apply only to new AutoShapes, and will not apply to new text boxes. If this method is called on a text box, the default formatting of that object will apply only to new text boxes, and will not apply to new AutoShapes.


## Example

This example adds a rectangle to the active publication, formats the rectangle's fill, applies the rectangle's formatting to the default shape, and then adds another smaller rectangle to the document. The second rectangle has the same fill as the first one.

```vb
With ActiveDocument.Pages(1).Shapes 
 
 With .AddShape(Type:=msoShapeRectangle, _ 
 Left:=5, Top:=5, Width:=80, Height:=60) 
 With .Fill 
 .ForeColor.RGB = RGB(0, 0, 255) 
 .BackColor.RGB = RGB(0, 204, 255) 
 .Patterned Pattern:=msoPatternHorizontalBrick 
 End With 
 .SetShapesDefaultProperties 
 End With 
 
 .AddShape Type:=msoShapeRectangle, _ 
 Left:=90, Top:=90, Width:=40, Height:=30 
 
End With 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]