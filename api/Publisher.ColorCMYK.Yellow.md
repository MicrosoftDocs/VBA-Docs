---
title: ColorCMYK.Yellow property (Publisher)
keywords: vbapb10.chm2621446
f1_keywords:
- vbapb10.chm2621446
ms.prod: publisher
api_name:
- Publisher.ColorCMYK.Yellow
ms.assetid: 2eaa27a4-a9bb-e18a-bd0e-c9cf07c567f5
ms.date: 06/06/2019
localization_priority: Normal
---


# ColorCMYK.Yellow property (Publisher)

Sets or returns a **Long** that represents the yellow component of a CMYK color. Value can be any number between 0 and 255. Read/write.


## Syntax

_expression_.**Yellow**

_expression_ A variable that represents a **[ColorCMYK](Publisher.ColorCMYK.md)** object.


## Return value

Long


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
 
 'Sets new shape to current shapes CMYK colors 
 shpStar.Fill.ForeColor.CMYK.SetCMYK _ 
 Cyan:=lngCyan, Magenta:=lngMagenta, _ 
 Yellow:=lngYellow, Black:=lngBlack 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]