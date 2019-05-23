---
title: Selection.HasChildShapeRange property (Word)
keywords: vbawd10.chm158663678
f1_keywords:
- vbawd10.chm158663678
ms.prod: word
api_name:
- Word.Selection.HasChildShapeRange
ms.assetid: 1917754f-6080-8303-533e-b62607b87d41
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.HasChildShapeRange property (Word)

 **True** if the selection contains child shapes. Read-only **Boolean**.


## Syntax

_expression_. `HasChildShapeRange`

 _expression_ An expression that returns a **[Selection](Word.Selection.md)** object.


## Example

This example creates a new document with a drawing canvas, populates the drawing canvas with shapes, and then, after checking that the shapes are child shapes, fills the child shapes with a pattern.


```vb
Sub ChildShapes() 
 Dim docNew As Document 
 Dim shpCanvas As Shape 
 
 'Create a new document with a drawing canvas and shapes 
 Set docNew = Documents.Add 
 Set shpCanvas = docNew.Shapes.AddCanvas( _ 
 Left:=100, Top:=100, Width:=200, Height:=200) 
 shpCanvas.CanvasItems.AddShape msoShapeRectangle, _ 
 Left:=0, Top:=0, Width:=100, Height:=100 
 shpCanvas.CanvasItems.AddShape msoShapeOval, _ 
 Left:=0, Top:=50, Width:=100, Height:=100 
 shpCanvas.CanvasItems.AddShape msoShapeDiamond, _ 
 Left:=0, Top:=100, Width:=100, Height:=100 
 
 'Select all shapes in the canvas 
 shpCanvas.CanvasItems.SelectAll 
 
 'Fill canvas child shapes with a pattern 
 If Selection.HasChildShapeRange = True Then 
 Selection.ChildShapeRange.Fill.Patterned msoPatternDivot 
 Else 
 MsgBox "This is not a range of child shapes." 
 End If 
 
End Sub
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]