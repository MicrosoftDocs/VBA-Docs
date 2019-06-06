---
title: FillFormat.OneColorGradient method (Publisher)
keywords: vbapb10.chm2359313
f1_keywords:
- vbapb10.chm2359313
ms.prod: publisher
api_name:
- Publisher.FillFormat.OneColorGradient
ms.assetid: e4ebf7c5-41af-8227-85de-10cc08ad9f91
ms.date: 06/07/2019
localization_priority: Normal
---


# FillFormat.OneColorGradient method (Publisher)

Sets the specified fill to a one-color gradient.


## Syntax

_expression_.**OneColorGradient** (_Style_, _Variant_, _Degree_)

_expression_ A variable that represents a **[FillFormat](publisher.fillformat.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Style_|Required| **[MsoGradientStyle](Office.MsoGradientStyle.md)** |The gradient style. Can be one of the **MsoGradientStyle** constants declared in the Microsoft Office type library.|
|_Variant_|Required| **Long**|The gradient variant. Can be a value from 1 to 4, corresponding to the four variants on the **Gradient** tab in the **Fill Effects** dialog box. If _Style_ is **msoGradientFromTitle** or **msoGradientFromCenter**, this argument can be either 1 or 2.|
|_Degree_|Required| **Single**|The gradient degree. Can be a value from 0.0 (dark) to 1.0 (light).|


## Example

This example adds a rectangle with a one-color gradient fill to the active publication.

```vb
With ActiveDocument.Pages(1).Shapes _ 
 .AddShape(Type:=msoShapeRectangle, _ 
 Left:=90, Top:=90, Width:=90, Height:=80).Fill 
 .ForeColor.RGB = RGB(0, 128, 128) 
 .OneColorGradient Style:=msoGradientHorizontal, _ 
 Variant:=1, Degree:=1 
End With 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]