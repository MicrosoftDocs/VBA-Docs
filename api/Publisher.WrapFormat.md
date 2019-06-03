---
title: WrapFormat object (Publisher)
keywords: vbapb10.chm851967
f1_keywords:
- vbapb10.chm851967
ms.prod: publisher
api_name:
- Publisher.WrapFormat
ms.assetid: b6f80d40-2043-6944-3ed8-f26635c7fa4d
ms.date: 06/04/2019
localization_priority: Normal
---


# WrapFormat object (Publisher)

Represents all the properties for wrapping text around a shape or shape range.
 
## Remarks

Use the **[Shape.TextWrap](Publisher.Shape.TextWrap.md)** property to return a **WrapFormat** object. 

## Example

The following example adds an oval to the active publication and specifies that publication text wrap around the left and right sides of the square that circumscribes the oval. There will be a 0.1-inch margin between the publication text and the top, bottom, left side, and right side of the square.

```vb
Sub SetTextWrapFormatProperties() 
 Dim shpOval As Shape 
 
 Set shpOval = ActiveDocument.Pages(1).Shapes.AddShape(Type:=msoShapeOval, _ 
 Left:=36, Top:=36, Width:=100, Height:=35) 
 With shpOval.TextWrap 
 .Type = pbWrapTypeSquare 
 .Side = pbWrapSideBoth 
 .DistanceAuto = msoFalse 
 .DistanceTop = InchesToPoints(0.1) 
 .DistanceBottom = InchesToPoints(0.1) 
 .DistanceLeft = InchesToPoints(0.1) 
 .DistanceRight = InchesToPoints(0.1) 
 End With 
End Sub
```


## Properties

- [Application](Publisher.WrapFormat.Application.md)
- [DistanceAuto](Publisher.WrapFormat.DistanceAuto.md)
- [DistanceBottom](Publisher.WrapFormat.DistanceBottom.md)
- [DistanceLeft](Publisher.WrapFormat.DistanceLeft.md)
- [DistanceRight](Publisher.WrapFormat.DistanceRight.md)
- [DistanceTop](Publisher.WrapFormat.DistanceTop.md)
- [Parent](Publisher.WrapFormat.Parent.md)
- [Side](Publisher.WrapFormat.Side.md)
- [Type](Publisher.WrapFormat.Type.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]