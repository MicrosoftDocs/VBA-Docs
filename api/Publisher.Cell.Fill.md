---
title: Cell.Fill property (Publisher)
keywords: vbapb10.chm5111817
f1_keywords:
- vbapb10.chm5111817
ms.prod: publisher
api_name:
- Publisher.Cell.Fill
ms.assetid: 3ff3fda8-aca7-534e-6a56-99d6a3d9e9e2
ms.date: 06/06/2019
localization_priority: Normal
---


# Cell.Fill property (Publisher)

Returns a **[FillFormat](Publisher.FillFormat.md)** object that represents the fill for the specified shape or table cell.


## Syntax

_expression_.**Fill**

_expression_ A variable that represents a **[Cell](Publisher.Cell.md)** object.


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