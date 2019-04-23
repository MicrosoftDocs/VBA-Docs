---
title: TextEffectFormat.PresetTextEffect property (Word)
keywords: vbawd10.chm164561004
f1_keywords:
- vbawd10.chm164561004
ms.prod: word
api_name:
- Word.TextEffectFormat.PresetTextEffect
ms.assetid: 86865b25-a30f-ef47-630f-b78ff1da28e3
ms.date: 06/08/2017
localization_priority: Normal
---


# TextEffectFormat.PresetTextEffect property (Word)

Returns or sets the style of the specified WordArt. The values for this property correspond to the formats in the  **WordArt Gallery** dialog box (**Insert** menu), numbered from left to right, top to bottom. Read/write **MsoPresetTextEffect**.


## Syntax

_expression_. `PresetTextEffect`

_expression_ Required. A variable that represents a '[TextEffectFormat](Word.TextEffectFormat.md)' object.


## Remarks

Setting the  **PresetTextEffect** property automatically sets many other formatting properties of the specified shape.


## Example

This example sets the style for all WordArt on myDocument to the first style listed in the WordArt Gallery dialog box.


```vb
Set myDocument = ActiveDocument 
For Each s In myDocument.Shapes 
 If s.Type = msoTextEffect Then 
 s.TextEffect.PresetTextEffect = msoTextEffect1 
 End If 
Next
```


## See also


[TextEffectFormat Object](Word.TextEffectFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]