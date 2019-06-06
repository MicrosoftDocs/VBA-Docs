---
title: ColorFormat.CMYK property (Publisher)
keywords: vbapb10.chm2555907
f1_keywords:
- vbapb10.chm2555907
ms.prod: publisher
api_name:
- Publisher.ColorFormat.CMYK
ms.assetid: 28d7ad65-c63c-3b11-3ecc-c77a1a586b84
ms.date: 06/06/2019
localization_priority: Normal
---


# ColorFormat.CMYK property (Publisher)

Returns a **[ColorCMYK](Publisher.ColorCMYK.md)** object that represents CMYK color properties.


## Syntax

_expression_.**CMYK**

_expression_ A variable that represents a **[ColorFormat](Publisher.ColorFormat.md)** object.


## Return value

ColorCMYK


## Example

This example creates two new shapes, and then sets the CMYK fill color for one shape and the CMYK values of the second shape to the same CMYK values.

```vb
Sub ReturnAndSetCMYK() 
 Dim lngCyan As Long 
 Dim lngMagenta As Long 
 Dim lngYellow As Long 
 Dim lngBlack As Long 
 Dim shpHeart As Shape 
 Dim shpStar As Shape 
 
 Set shpHeart = ActiveDocument.Pages(1).Shapes.AddShape _ 
 (Type:=msoShapeHeart, Left:=100, _ 
 Top:=100, Width:=100, Height:=100) 
 Set shpStar = ActiveDocument.Pages(1).Shapes.AddShape _ 
 (Type:=msoShape5pointStar, Left:=200, _ 
 Top:=100, Width:=150, Height:=150) 
 
 With shpHeart.Fill.ForeColor.CMYK 
 .SetCMYK 10, 80, 200, 30 
 lngCyan = .Cyan 
 lngMagenta = .Magenta 
 lngYellow = .Yellow 
 lngBlack = .Black 
 End With 
 
 'Sets new shape to current shape's CMYK colors 
 shpStar.Fill.ForeColor.CMYK.SetCMYK _ 
 Cyan:=lngCyan, Magenta:=lngMagenta, _ 
 Yellow:=lngYellow, Black:=lngBlack 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]