---
title: LineFormat.Style property (Publisher)
keywords: vbapb10.chm3408144
f1_keywords:
- vbapb10.chm3408144
ms.prod: publisher
api_name:
- Publisher.LineFormat.Style
ms.assetid: 3826eb43-b90e-e24b-31d5-8d9eddd3ed4e
ms.date: 06/08/2019
localization_priority: Normal
---


# LineFormat.Style property (Publisher)

Returns or sets an **[MsoLineStyle](office.msolinestyle.md)** constant that represents the style of line to apply to a shape or border. Read/write.


## Syntax

_expression_.**Style**

_expression_ A variable that represents a **[LineFormat](Publisher.LineFormat.md)** object.


## Return value

MsoLineStyle


## Remarks

The **Style** property value can be one of the **MsoLineStyle** constants declared in the Microsoft Office type library.

## Example

This example adds a new shape and sets the line properties for the shape.

```vb
Sub SetLineStyle() 
 With ActiveDocument.Pages(1).Shapes.AddShape(msoShapeRectangle, _ 
 Left:=72, Top:=140, Width:=200, Height:=100) 
 .Rotation = 120 
 With .Line 
 .Weight = 5 
 .DashStyle = msoLineDashDotDot 
 .Style = msoLineThickBetweenThin 
 End With 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]