---
title: Shape.CanvasItems property (Word)
keywords: vbawd10.chm161480842
f1_keywords:
- vbawd10.chm161480842
ms.prod: word
api_name:
- Word.Shape.CanvasItems
ms.assetid: 2dfe33c7-1487-6074-9135-2d3220e11691
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.CanvasItems property (Word)

Returns a  **[CanvasShapes](Word.CanvasShapes.md)** object that represents a collection of shapes in a drawing canvas.


## Syntax

 _expression_. `CanvasItems`

 _expression_ Required. A variable that represents a '[Shape](Word.Shape.md)' object.


## Example

This example creates a new drawing canvas in the active document and adds a circle to the canvas.


```vb
Sub NewCanvasShape() 
 Dim shpCanvas As Shape 
 Set shpCanvas = ActiveDocument.Shapes.AddCanvas( _ 
 Left:=100, Top:=75, Width:=150, Height:=200) 
 shpCanvas.CanvasItems.AddShape _ 
 Type:=msoShapeOval, Top:=25, _ 
 Left:=25, Width:=150, Height:=150 
End Sub
```


## See also


[Shape Object](Word.Shape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]