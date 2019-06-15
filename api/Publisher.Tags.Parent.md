---
title: Tags.Parent property (Publisher)
keywords: vbapb10.chm4653058
f1_keywords:
- vbapb10.chm4653058
ms.prod: publisher
api_name:
- Publisher.Tags.Parent
ms.assetid: 8d164438-f15f-5232-92b4-d293b62b15da
ms.date: 06/15/2019
localization_priority: Normal
---


# Tags.Parent property (Publisher)

Returns an object that represents the parent object of the specified object. For example, for a **[TextFrame](Publisher.TextFrame.md)** object, returns a **[Shape](Publisher.Shape.md)** object representing the parent shape of the text frame. Read-only.


## Syntax

_expression_.**Parent**

_expression_ A variable that represents a **[Tags](Publisher.Tags.md)** object.


## Example

This example accesses the parent object of the selected shape, and then adds a new shape to it and sets the fill for the new shape.

```vb
Sub ParentObject() 
 Dim shp As Shape 
 Dim pg As Page 
 
 Set pg = Selection.ShapeRange(1).Parent 
 Set shp = pg.Shapes.AddShape(Type:=msoShape5pointStar, _ 
 Left:=72, Top:=72, Width:=72, Height:=72) 
 
 shp.Fill.ForeColor.RGB = RGB(Red:=180, Green:=180, Blue:=180) 
End Sub
```

<br/>

This example returns the parent object of a text frame, which is the first shape in the active publication, and then fills the shape with a pattern.

```vb
Sub ParentShape() 
 Dim shpParent As Shape 
 Set shpParent = ActiveDocument.Pages(1).Shapes(1).TextFrame.Parent 
 shpParent.Fill.Patterned Pattern:=msoPatternSphere 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]