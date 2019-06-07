---
title: FillFormat.GradientStyle property (Publisher)
keywords: vbapb10.chm2359556
f1_keywords:
- vbapb10.chm2359556
ms.prod: publisher
api_name:
- Publisher.FillFormat.GradientStyle
ms.assetid: 38a38de1-4ed3-7919-421f-474b0b5d7b2f
ms.date: 06/07/2019
localization_priority: Normal
---


# FillFormat.GradientStyle property (Publisher)

Returns an **[MsoGradientStyle](Office.MsoGradientStyle.md)** constant indicating the gradient style for the specified fill. Read-only.


## Syntax

_expression_.**GradientStyle**

_expression_ A variable that represents a **[FillFormat](publisher.fillformat.md)** object.


## Return value

MsoGradientStyle


## Remarks

Use the **[OneColorGradient](Publisher.FillFormat.OneColorGradient.md)**, **[PresetGradient](Publisher.FillFormat.PresetGradient.md)**, or **[TwoColorGradient](Publisher.FillFormat.TwoColorGradient.md)** method to set the gradient style for the fill.

Attempting to return this property for a fill that doesn't have a gradient generates an error. Use the **[Type](Publisher.FillFormat.Type.md)** property to determine whether the fill has a gradient.

The **GradientStyle** property value can be one of the **MsoGradientStyle** constants declared in the Microsoft Office type library.


## Example

This example adds a rectangle to the active publication and sets its fill gradient style to match that of the shape named rect1. For the example to work, rect1 must have a gradient fill.

```vb
Dim intStyle As Integer 
 
With ActiveDocument.Pages(1).Shapes 
 ' Store gradient style of rect1. 
 intStyle = .Item("rect1").Fill.GradientStyle 
 ' Add new rectangle. 
 With .AddShape(Type:=msoShapeRectangle, _ 
 Left:=0, Top:=0, Width:=40, Height:=80).Fill 
 ' Set color and gradient of new rectangle. 
 .ForeColor.RGB = RGB(128, 0, 0) 
 .OneColorGradient Style:=intStyle, _ 
 Variant:=1, Degree:=1 
 End With 
End With 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]