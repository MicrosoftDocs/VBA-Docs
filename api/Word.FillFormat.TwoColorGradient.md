---
title: FillFormat.TwoColorGradient method (Word)
keywords: vbawd10.chm164102160
f1_keywords:
- vbawd10.chm164102160
ms.prod: word
api_name:
- Word.FillFormat.TwoColorGradient
ms.assetid: 38a8a57c-0f5f-e54a-999e-87e0977196b8
ms.date: 06/08/2017
localization_priority: Normal
---


# FillFormat.TwoColorGradient method (Word)

Sets the specified fill to a two-color gradient.


## Syntax

_expression_.**TwoColorGradient** (_Style_, _Variant_)

_expression_ Required. A variable that represents a **[FillFormat](word.fillformat.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Style_|Required| **[MsoGradientStyle](Office.MsoGradientStyle.md)**|The gradient style. Can be any **MsoGradientStyle** constant except **msoGradientFromTitle** which applies only to Microsoft PowerPoint.|
| _Variant_|Required| **Long**|The gradient variant. Can be a value from 1 to 4, corresponding to the four variants on the **Gradient** tab in the **Fill Effects** dialog box. If _Style_ is **msoGradientFromCenter**, this argument can be either 1 or 2.|


## Example

This example adds a rectangle with a two-color gradient fill to the active document and sets the background and foreground color for the fill.

```vb
With ActiveDocument.Shapes.AddShape(msoShapeRectangle, _ 
 0, 0, 40, 80).Fill 
 .ForeColor.RGB = RGB(128, 0, 0) 
 .BackColor.RGB = RGB(0, 170, 170) 
 .TwoColorGradient msoGradientHorizontal, 1 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]