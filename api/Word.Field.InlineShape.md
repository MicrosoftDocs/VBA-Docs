---
title: Field.InlineShape property (Word)
keywords: vbawd10.chm154075148
f1_keywords:
- vbawd10.chm154075148
ms.prod: word
api_name:
- Word.Field.InlineShape
ms.assetid: 2fbaa2a5-3c31-e7ff-45db-044c62cde951
ms.date: 06/08/2017
localization_priority: Normal
---


# Field.InlineShape property (Word)

Returns an  **[InlineShape](Word.InlineShape.md)** object that represents the picture, OLE object, or ActiveX control that is the result of an INCLUDEPICTURE or EMBED field.


## Syntax

_expression_. `InlineShape`

 _expression_ An expression that returns a '[Field](Word.Field.md)' object.


## Remarks

An  **InlineShape** object is treated like a character and is positioned as a character within a line of text.


## Example

This example returns the width of the inline shape associated with the first field in the active document. For this example to work, the field must be an INCLUDEPICTURE field.


```vb
If ActiveDocument.Fields(1).Type = wdFieldIncludePicture Then 
 MsgBox ActiveDocument.Fields(1).InlineShape.Width 
End If
```


## See also


[Field Object](Word.Field.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]