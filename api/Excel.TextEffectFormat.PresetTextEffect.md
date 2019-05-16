---
title: TextEffectFormat.PresetTextEffect property (Excel)
keywords: vbaxl10.chm118010
f1_keywords:
- vbaxl10.chm118010
ms.prod: excel
api_name:
- Excel.TextEffectFormat.PresetTextEffect
ms.assetid: 13ff8b1a-d12e-47c1-6f82-0b3b9b5a7bf4
ms.date: 05/17/2019
localization_priority: Normal
---


# TextEffectFormat.PresetTextEffect property (Excel)

Returns or sets the style of the specified WordArt. Read/write **[MsoPresetTextEffect](office.msopresettexteffect.md)**.


## Syntax

_expression_.**PresetTextEffect**

_expression_ A variable that represents a **[TextEffectFormat](Excel.TextEffectFormat.md)** object.


## Remarks

The values for this property correspond to the formats in the **WordArt Gallery** dialog box (numbered from left to right, top to bottom).

Setting the **PresetTextEffect** property automatically sets many other formatting properties of the specified shape.


## Example

This example sets the style for all WordArt on _myDocument_ to the first style listed in the **WordArt Gallery** dialog box.

```vb
Set myDocument = Worksheets(1) 
For Each s In myDocument.Shapes 
 If s.Type = msoTextEffect Then 
 s.TextEffect.PresetTextEffect = msoTextEffect1 
 End If 
Next
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]