---
title: TextEffectFormat object (Publisher)
keywords: vbapb10.chm3801087
f1_keywords:
- vbapb10.chm3801087
ms.prod: publisher
api_name:
- Publisher.TextEffectFormat
ms.assetid: 672d0ef0-cbcd-05ef-9aa5-b986c7b045ac
ms.date: 06/01/2019
localization_priority: Normal
---


# TextEffectFormat object (Publisher)

Contains properties and methods that apply to WordArt objects.
 
## Remarks

Use the **[TextEffect](publisher.shape.texteffect.md)** property of the **Shape** or **[ShapeRange](publisher.shaperange.texteffect.md)** object to return a **TextEffectFormat** object. 

## Example

The following example sets the font name and formatting for shape one on the first page of the active publication. For this example to work, shape one must be a WordArt object.

```vb
Sub FormatWordArt() 
 With ActiveDocument.Pages(1).Shapes(1).TextEffect 
 .FontName = "Courier New" 
 .FontBold = MsoTrue 
 .FontItalic = MsoTrue 
 End With 
End Sub
```


## Methods

- [ToggleVerticalText](Publisher.TextEffectFormat.ToggleVerticalText.md)

## Properties

- [Alignment](Publisher.TextEffectFormat.Alignment.md)
- [Application](Publisher.TextEffectFormat.Application.md)
- [FontBold](Publisher.TextEffectFormat.FontBold.md)
- [FontItalic](Publisher.TextEffectFormat.FontItalic.md)
- [FontName](Publisher.TextEffectFormat.FontName.md)
- [FontSize](Publisher.TextEffectFormat.FontSize.md)
- [KernedPairs](Publisher.TextEffectFormat.KernedPairs.md)
- [NormalizedHeight](Publisher.TextEffectFormat.NormalizedHeight.md)
- [Parent](Publisher.TextEffectFormat.Parent.md)
- [PresetShape](Publisher.TextEffectFormat.PresetShape.md)
- [PresetTextEffect](Publisher.TextEffectFormat.PresetTextEffect.md)
- [PresetWordArt](Publisher.TextEffectFormat.PresetWordArt.md)
- [RotatedChars](Publisher.TextEffectFormat.RotatedChars.md)
- [Text](Publisher.TextEffectFormat.Text.md)
- [Tracking](Publisher.TextEffectFormat.Tracking.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]