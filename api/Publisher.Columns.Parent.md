---
title: Columns.Parent Property (Publisher)
keywords: vbapb10.chm5046275
f1_keywords:
- vbapb10.chm5046275
ms.prod: publisher
api_name:
- Publisher.Columns.Parent
ms.assetid: 735fc7b8-ca76-0573-d1fa-2b5dff156f6e
ms.date: 06/08/2017
---


# Columns.Parent Property (Publisher)

Returns an object that represents the parent object of the specified object. For example, for a  **[TextFrame](Publisher.TextFrame.md)** object, returns a **[Shape](Publisher.Shape.md)** object representing the parent shape of the text frame. Read-only.


## Syntax

 _expression_. **Parent**

 _expression_ A variable that represents a  **Columns** object.


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

This example returns the parent object of a text frame, which is the first shape in the active publication, and then fills the shape with a pattern.




```vb
Sub ParentShape() 
 Dim shpParent As Shape 
 Set shpParent = ActiveDocument.Pages(1).Shapes(1).TextFrame.Parent 
 shpParent.Fill.Patterned Pattern:=msoPatternSphere 
End Sub
```


