---
title: ColorCMYK.SetCMYK method (Publisher)
keywords: vbapb10.chm2621447
f1_keywords:
- vbapb10.chm2621447
ms.prod: publisher
api_name:
- Publisher.ColorCMYK.SetCMYK
ms.assetid: 9c7ec18b-73e9-66bc-57f4-cd6d62817630
ms.date: 06/06/2019
localization_priority: Normal
---


# ColorCMYK.SetCMYK method (Publisher)

Sets a cyan-magenta-yellow-black (CMYK) color value.


## Syntax

_expression_.**SetCMYK** (_Cyan_, _Magenta_, _Yellow_, _Black_)

_expression_ A variable that represents a **[ColorCMYK](Publisher.ColorCMYK.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Cyan_|Required| **Long**|A number that represents the cyan component of the color. Value can be any number between 0 and 255.|
|_Magenta_|Required| **Long**|A number that represents the magenta component of the color. Value can be any number between 0 and 255.|
|_Yellow_|Required| **Long**|A number that represents the yellow component of the color. Value can be any number between 0 and 255.|
|_Black_|Required| **Long**|A number that represents the black component of the color. Value can be any number between 0 and 255.|

## Example

This example sets the CMYK color for the specified shape.

```vb
Sub SetCMYKColor() 
 Dim shpStar As Shape 
 
 Set shpStar = ActiveDocument.Pages(1).Shapes _ 
 .AddShape(Type:=msoShape5pointStar, Left:=72, _ 
 Top:=72, Width:=150, Height:=150) 
 shpStar.Fill.ForeColor.CMYK.SetCMYK Cyan:=0, _ 
 Magenta:=255, Yellow:=255, Black:=50 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]