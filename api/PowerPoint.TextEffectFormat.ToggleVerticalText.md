---
title: TextEffectFormat.ToggleVerticalText method (PowerPoint)
keywords: vbapp10.chm556002
f1_keywords:
- vbapp10.chm556002
ms.prod: powerpoint
api_name:
- PowerPoint.TextEffectFormat.ToggleVerticalText
ms.assetid: f9b71bae-4432-c4bd-4b47-1294520e33d1
ms.date: 06/08/2017
localization_priority: Normal
---


# TextEffectFormat.ToggleVerticalText method (PowerPoint)

Switches the text flow in the specified WordArt from horizontal to vertical, or vice versa.


## Syntax

_expression_. `ToggleVerticalText`

_expression_ A variable that represents a [TextEffectFormat](PowerPoint.TextEffectFormat.md) object.


## Remarks

Using the  **ToggleVerticalText** method swaps the values of the **Width** and **Height** properties of the **Shape** object that represents the WordArt and leaves the **Left** and **Top** properties unchanged.

The **[Flip](PowerPoint.Shape.Flip.md)** method and **[Rotation](PowerPoint.Shape.Rotation.md)** property of the **[Shape](PowerPoint.Shape.md)** object and the **[RotatedChars](PowerPoint.TextEffectFormat.RotatedChars.md)** property and **ToggleVerticalText** method of the **TextEffectFormat** object all affect the character orientation and the direction of text flow in a **Shape** object that represents WordArt. You may have to experiment to find out how to combine the effects of these properties and methods to get the result you want.


## Example

This example adds WordArt that contains the text "Test" to _myDocument_, and switches from horizontal text flow (the default for the specified WordArt style,  **msoTextEffect1**) to vertical text flow.


```vb
Set myDocument = ActivePresentation.Slides(1)

Set newWordArt = myDocument.Shapes.AddTextEffect _
    (PresetTextEffect:=msoTextEffect1, Text:="Test", _
    FontName:="Arial Black", FontSize:=36, _
    FontBold:=False, FontItalic:=False, Left:=100, Top:=100)
newWordArt.TextEffect.ToggleVerticalText
```


## See also


[TextEffectFormat Object](PowerPoint.TextEffectFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]