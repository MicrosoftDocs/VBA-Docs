---
title: ShapeRange.Fill property (Publisher)
keywords: vbapb10.chm2293815
f1_keywords:
- vbapb10.chm2293815
ms.prod: publisher
api_name:
- Publisher.ShapeRange.Fill
ms.assetid: cdff2b6f-52f5-3ab3-c57a-4647888cd96f
ms.date: 06/14/2019
localization_priority: Normal
---


# ShapeRange.Fill property (Publisher)

Returns a **[FillFormat](Publisher.FillFormat.md)** object representing the fill for the specified shape or table cell.


## Syntax

_expression_.**Fill**

_expression_ A variable that represents a **[ShapeRange](Publisher.ShapeRange.md)** object.


## Example

This example creates a new AutoShape object and fills the shape with green.

```vb
Sub NewShapeItem() 
 
 Dim shpHeart As Shape 
 
 Set shpHeart = ThisDocument.MasterPages.Item(1).Shapes _ 
 .AddShape(Type:=msoShapeHeart, Left:=40, Top:=80, _ 
 Width:=50, Height:=50) 
 shpHeart.Fill.ForeColor.RGB = RGB(Red:=0, Green:=255, Blue:=0) 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]