---
title: TextEffectFormat.PresetShape property (Excel)
keywords: vbaxl10.chm118009
f1_keywords:
- vbaxl10.chm118009
ms.prod: excel
api_name:
- Excel.TextEffectFormat.PresetShape
ms.assetid: d85bcdf6-0ad4-4a3d-ed3a-40a96a4b63a0
ms.date: 05/17/2019
localization_priority: Normal
---


# TextEffectFormat.PresetShape property (Excel)

Returns or sets the shape of the specified WordArt. Read/write **[MsoPresetTextEffectShape](office.msopresettexteffectshape.md)**.


## Syntax

_expression_.**PresetShape**

_expression_ A variable that represents a **[TextEffectFormat](Excel.TextEffectFormat.md)** object.


## Remarks

Setting the **[PresetTextEffect](Excel.TextEffectFormat.PresetTextEffect.md)** property automatically sets the **PresetShape** property.


## Example

This example sets the shape of all WordArt on _myDocument_ to a chevron whose center points down.

```vb
Set myDocument = Worksheets(1) 
For Each s In myDocument.Shapes 
 If s.Type = msoTextEffect Then 
 s.TextEffect.PresetShape = msoTextEffectShapeChevronDown 
 End If 
Next
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]