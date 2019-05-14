---
title: InlineShape.TextEffect property (Word)
keywords: vbawd10.chm162005112
f1_keywords:
- vbawd10.chm162005112
ms.prod: word
api_name:
- Word.InlineShape.TextEffect
ms.assetid: 349563af-6a14-a8d9-c0a4-829910d7dc2c
ms.date: 06/08/2017
localization_priority: Normal
---


# InlineShape.TextEffect property (Word)

Returns a  **TextEffectFormat** object that contains text-effect formatting properties for the specified inline shape. Read-only.


## Syntax

_expression_.**TextEffect**

_expression_ A variable that represents an '[InlineShape](Word.InlineShape.md)' object.


## Example

This example sets the font style to bold for shape three on _myDocument_ if the shape is WordArt.


```vb
Set myDocument = ActiveDocument 
With myDocument.Shapes(3) 
 If .Type = msoTextEffect Then 
 .TextEffect.FontBold = True 
 End If 
End With
```


## See also


[InlineShape Object](Word.InlineShape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]