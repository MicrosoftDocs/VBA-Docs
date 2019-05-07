---
title: Envelope.DefaultHeight property (Word)
keywords: vbawd10.chm152567814
f1_keywords:
- vbawd10.chm152567814
ms.prod: word
api_name:
- Word.Envelope.DefaultHeight
ms.assetid: 4c13a3b2-4236-defa-3682-ccef1700901f
ms.date: 06/08/2017
localization_priority: Normal
---


# Envelope.DefaultHeight property (Word)

Returns or sets the default envelope height, in points. Read/write  **Single**.


## Syntax

_expression_. `DefaultHeight`

_expression_ A variable that represents a '[Envelope](Word.Envelope.md)' object.


## Remarks

If you set either the  **DefaultHeight** or **[DefaultWidth](Word.Envelope.DefaultWidth.md)** property, the envelope size is automatically changed to **Custom Size** in the **Envelope Options** dialog box (**Tools** menu). Use the **[DefaultSize](Word.Envelope.DefaultSize.md)** property to set the default size to a predefined size.


## Example

This example sets the default envelope size to 4.5 inches by 7.5 inches.


```vb
With ActiveDocument.Envelope 
 .DefaultHeight = InchesToPoints(4.5) 
 .DefaultWidth = InchesToPoints(7.5) 
End With
```


## See also


[Envelope Object](Word.Envelope.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]