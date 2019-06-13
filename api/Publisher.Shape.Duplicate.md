---
title: Shape.Duplicate method (Publisher)
keywords: vbapb10.chm2228244
f1_keywords:
- vbapb10.chm2228244
ms.prod: publisher
api_name:
- Publisher.Shape.Duplicate
ms.assetid: 9f35a496-5312-bff1-a31e-05baaaf69e92
ms.date: 06/13/2019
localization_priority: Normal
---


# Shape.Duplicate method (Publisher)

Creates a duplicate of the specified **Shape** or **[ShapeRange](Publisher.ShapeRange.md)** object, adds the new shape or range of shapes to the **[Shapes](Publisher.Shapes.md)** collection immediately after the shape or range of shapes specified originally, and then returns the new **Shape** or **ShapeRange** object.


## Syntax

_expression_.**Duplicate**

_expression_ A variable that represents a **[Shape](Publisher.Shape.md)** object.


## Return value

Shape


## Example

This example adds a new, blank page at the end of the active publication, adds a diamond shape to the new page, duplicates the diamond, and then sets properties for the duplicate. The first diamond will have the default fill color for the active color scheme; the second diamond will be offset from the first one and will have the first accent color for the active color scheme.

```vb
Dim pgTemp As Page 
Dim shpTemp As Shape 
 
Set pgTemp = ActiveDocument.Pages.Add(Count:=1, After:=1) 
Set shpTemp = pgTemp.Shapes _ 
 .AddShape(Type:=msoShapeDiamond, _ 
 Left:=10, Top:=10, Width:=250, Height:=350) 
 
With shpTemp.Duplicate 
 .Left = 150 
 .Fill.ForeColor.SchemeColor = pbSchemeColorAccent1 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]