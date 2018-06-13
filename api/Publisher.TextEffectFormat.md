---
title: TextEffectFormat Object (Publisher)
keywords: vbapb10.chm3801087
f1_keywords:
- vbapb10.chm3801087
ms.prod: publisher
api_name:
- Publisher.TextEffectFormat
ms.assetid: 672d0ef0-cbcd-05ef-9aa5-b986c7b045ac
ms.date: 06/08/2017
---


# TextEffectFormat Object (Publisher)

Contains properties and methods that apply to WordArt objects.
 


## Example

Use the  **TextEffect** property to return a **TextEffectFormat** object. The following example sets the font name and formatting for shape one on the first page of the active publication. For this example to work, shape one must be a WordArt object.
 

 

```
Sub FormatWordArt() 
 With ActiveDocument.Pages(1).Shapes(1).TextEffect 
 .FontName = "Courier New" 
 .FontBold = MsoTrue 
 .FontItalic = MsoTrue 
 End With 
End Sub
```


## Methods



|**Name**|
|:-----|
|[ToggleVerticalText](Publisher.TextEffectFormat.ToggleVerticalText.md)|

## Properties



|**Name**|
|:-----|
|[Alignment](Publisher.TextEffectFormat.Alignment.md)|
|[Application](Publisher.TextEffectFormat.Application.md)|
|[FontBold](Publisher.TextEffectFormat.FontBold.md)|
|[FontItalic](Publisher.TextEffectFormat.FontItalic.md)|
|[FontName](Publisher.TextEffectFormat.FontName.md)|
|[FontSize](Publisher.TextEffectFormat.FontSize.md)|
|[KernedPairs](Publisher.TextEffectFormat.KernedPairs.md)|
|[NormalizedHeight](Publisher.TextEffectFormat.NormalizedHeight.md)|
|[Parent](Publisher.TextEffectFormat.Parent.md)|
|[PresetShape](Publisher.TextEffectFormat.PresetShape.md)|
|[PresetTextEffect](Publisher.TextEffectFormat.PresetTextEffect.md)|
|[PresetWordArt](Publisher.TextEffectFormat.PresetWordArt.md)|
|[RotatedChars](Publisher.TextEffectFormat.RotatedChars.md)|
|[Text](Publisher.TextEffectFormat.Text.md)|
|[Tracking](Publisher.TextEffectFormat.Tracking.md)|

