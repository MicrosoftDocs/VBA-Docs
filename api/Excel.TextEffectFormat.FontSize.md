---
title: TextEffectFormat.FontSize property (Excel)
keywords: vbaxl10.chm118006
f1_keywords:
- vbaxl10.chm118006
ms.prod: excel
api_name:
- Excel.TextEffectFormat.FontSize
ms.assetid: b78fa323-4fcb-c12a-4166-f1689d9f0a93
ms.date: 05/17/2019
localization_priority: Normal
---


# TextEffectFormat.FontSize property (Excel)

Returns or sets the font size for the specified WordArt, in [points](../language/glossary/vbe-glossary.md#point). Read/write **Single**.


## Syntax

_expression_.**FontSize**

_expression_ A variable that represents a **[TextEffectFormat](Excel.TextEffectFormat.md)** object.


## Example

This example sets the font size to 16 points for the shape named WordArt 4 in _myDocument_.

```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes("WordArt 4").TextEffect.FontSize = 16
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]