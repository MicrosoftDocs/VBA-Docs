---
title: TextEffectFormat object (Word)
keywords: vbawd10.chm2511
f1_keywords:
- vbawd10.chm2511
ms.prod: word
api_name:
- Word.TextEffectFormat
ms.assetid: b274e5be-ed5b-7d63-aa4b-1d67b63e7c0b
ms.date: 06/08/2017
localization_priority: Normal
---


# TextEffectFormat object (Word)

Contains properties and methods that apply to WordArt objects.


## Remarks

Use the **TextEffect** property to return a **TextEffectFormat** object. The following example sets the font name and formatting for shape one on the active document. For this example to work, shape one must be a WordArt object.


```vb
With ActiveDocument.Shapes(1).TextEffect 
 .FontName = "Courier New" 
 .FontBold = True 
 .FontItalic = True 
End With
```


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]