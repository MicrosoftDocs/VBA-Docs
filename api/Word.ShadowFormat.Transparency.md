---
title: ShadowFormat.Transparency property (Word)
keywords: vbawd10.chm164364392
f1_keywords:
- vbawd10.chm164364392
ms.prod: word
api_name:
- Word.ShadowFormat.Transparency
ms.assetid: 0bcfcbd7-ffde-4dc0-628a-203c322cf573
ms.date: 06/08/2017
localization_priority: Normal
---


# ShadowFormat.Transparency property (Word)

Returns or sets the degree of transparency of the specified shadow as a value between 0.0 (opaque) and 1.0 (clear). Read/write  **Single**.


## Syntax

_expression_.**Transparency**

_expression_ Required. A variable that represents a **[ShadowFormat](Word.ShadowFormat.md)** object.


## Example

This example sets the shadow of shape three in the active document to semitransparent red. If the shape doesn't already have a shadow, this example adds one to it.


```vb
With ActiveDocument.Shapes(3).Shadow 
 .Visible = True 
 .ForeColor.RGB = RGB(255, 0, 0) 
 .Transparency = 0.5 
End With
```


## See also


[ShadowFormat Object](Word.ShadowFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]