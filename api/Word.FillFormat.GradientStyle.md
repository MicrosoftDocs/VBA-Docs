---
title: FillFormat.GradientStyle property (Word)
keywords: vbawd10.chm164102248
f1_keywords:
- vbawd10.chm164102248
ms.prod: word
api_name:
- Word.FillFormat.GradientStyle
ms.assetid: f5342153-6160-c1cd-c02f-708c7c08a902
ms.date: 06/08/2017
localization_priority: Normal
---


# FillFormat.GradientStyle property (Word)

Returns the gradient style for the specified fill. Read-only  **MsoGradientStyle**.


## Syntax

_expression_.**GradientStyle**

 _expression_ An expression that represents a **[FillFormat](word.fillformat.md)** object.


## Remarks

This property is read-only. Use the **[OneColorGradient](Word.FillFormat.OneColorGradient.md)** or **[TwoColorGradient](Word.FillFormat.TwoColorGradient.md)** method to set the gradient style for the fill.

Attempting to return this property for a fill that doesn't have a gradient generates an error. Use the **[Type](Word.FillFormat.Type.md)** property to determine whether the fill has a gradient.


## Example

This example adds a rectangle to the active document and sets its fill gradient style to match that of the shape named "rect1." For the example to work, rect1 must have a gradient fill.


```vb
Dim docActive As Document 
Dim lngGradient As Long 
 
Set docActive = ActiveDocument 
With docActive.Shapes 
 lngGradient = .Item("rect1").Fill.GradientStyle 
 With .AddShape(msoShapeRectangle, 0, 0, 40, 80).Fill 
 .ForeColor.RGB = RGB(128, 0, 0) 
 .OneColorGradient lngGradient, 1, 1 
 End With 
End With
```


## See also


[FillFormat Object](Word.FillFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]