---
title: LineFormat.InsetPen property (Excel)
ms.prod: excel
api_name:
- Excel.LineFormat.InsetPen
ms.assetid: 7a9999ad-b3a5-bae5-e068-8d85cab5ecb5
ms.date: 04/30/2019
localization_priority: Normal
---


# LineFormat.InsetPen property (Excel)

Returns or sets whether lines are drawn inside the specified shape's boundaries. Read/write.


## Syntax

_expression_.**InsetPen**

_expression_ A variable that represents a **[LineFormat](Excel.LineFormat.md)** object.


## Return value

**[MsoTriState](Office.MsoTriState.md)**


## Remarks

**msoTrue** (-1) if lines are drawn inside the shape's boundaries; otherwise, **msoFalse** (0).


## Example

The following code example adds two rectangles to the active worksheet, the first with its lines drawn inside its boundaries, and the second with its lines drawn on its boundaries.

```vb
Dim shpNew As Shape 
 
With ActiveSheet.Shapes 
 Set shpNew = .AddShape(Type:=msoShapeRectangle, _ 
 Left:=200, Top:=150, Width:=150, Height:=100) 
 With shpNew.Line 
 .Weight = 24 
 .InsetPen = msoTrue 
 End With 
 
 Set shpNew = .AddShape(Type:=msoShapeRectangle, _ 
 Left:=200, Top:=300, Width:=150, Height:=100) 
 With shpNew.Line 
 .Weight = 24 
 .InsetPen = msoFalse 
 End With 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]