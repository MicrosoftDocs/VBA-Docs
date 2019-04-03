---
title: TextEffectFormat object (Excel)
keywords: vbaxl10.chm118000
f1_keywords:
- vbaxl10.chm118000
ms.prod: excel
api_name:
- Excel.TextEffectFormat
ms.assetid: 7fe03721-6a45-569e-add4-fc8849c99535
ms.date: 04/02/2019
localization_priority: Normal
---


# TextEffectFormat object (Excel)

Contains properties and methods that apply to WordArt objects.


## Remarks

Use the **[TextEffect](Excel.Shape.TextEffect.md)** property of the **Shape** object to return a **TextEffectFormat** object.


## Example

The following example sets the font name and formatting for shape one on _myDocument_. For this example to work, shape one must be a WordArt object.

```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(1).TextEffect 
 .FontName = "Courier New" 
 .FontBold = True 
 .FontItalic = True 
End With 

```

## Methods

- [ToggleVerticalText](Excel.TextEffectFormat.ToggleVerticalText.md)

## Properties

- [Alignment](Excel.TextEffectFormat.Alignment.md)
- [Application](Excel.TextEffectFormat.Application.md)
- [Creator](Excel.TextEffectFormat.Creator.md)
- [FontBold](Excel.TextEffectFormat.FontBold.md)
- [FontItalic](Excel.TextEffectFormat.FontItalic.md)
- [FontName](Excel.TextEffectFormat.FontName.md)
- [FontSize](Excel.TextEffectFormat.FontSize.md)
- [KernedPairs](Excel.TextEffectFormat.KernedPairs.md)
- [NormalizedHeight](Excel.TextEffectFormat.NormalizedHeight.md)
- [Parent](Excel.TextEffectFormat.Parent.md)
- [PresetShape](Excel.TextEffectFormat.PresetShape.md)
- [PresetTextEffect](Excel.TextEffectFormat.PresetTextEffect.md)
- [RotatedChars](Excel.TextEffectFormat.RotatedChars.md)
- [Text](Excel.TextEffectFormat.Text.md)
- [Tracking](Excel.TextEffectFormat.Tracking.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]