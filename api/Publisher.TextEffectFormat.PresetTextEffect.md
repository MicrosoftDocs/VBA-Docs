---
title: TextEffectFormat.PresetTextEffect property (Publisher)
keywords: vbapb10.chm3735816
f1_keywords:
- vbapb10.chm3735816
ms.prod: publisher
api_name:
- Publisher.TextEffectFormat.PresetTextEffect
ms.assetid: d7ef0995-4560-fea0-b98d-03c8e0c8e65e
ms.date: 06/15/2019
localization_priority: Normal
---


# TextEffectFormat.PresetTextEffect property (Publisher)

Returns or sets an **[MsoPresetTextEffect](Office.MsoPresetTextEffect.md)** constant that represents the style of the specified WordArt. The values for this property correspond to the formats in the **WordArt Gallery** dialog box, numbered from left to right, top to bottom. Read/write.


## Syntax

_expression_.**PresetTextEffect**

_expression_ A variable that represents a **[TextEffectFormat](Publisher.TextEffectFormat.md)** object.


## Return value

MsoPresetTextEffect


## Remarks

The **PresetTextEffect** property value can be one of the **MsoPresetTextEffect** constants declared in the Microsoft Office type library.


## Example

This example sets the text effect style for the first shape on the first page of the active publication. This example assumes that there is at least one shape on the first page of the active publication.

```vb
Sub ChangeTextEffect() 
 With ActiveDocument.Pages(1).Shapes(1) 
 If .Type = msoTextEffect Then 
 .TextEffect.PresetTextEffect = msoTextEffect1 
 End If 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]