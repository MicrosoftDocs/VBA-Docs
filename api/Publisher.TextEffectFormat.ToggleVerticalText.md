---
title: TextEffectFormat.ToggleVerticalText Method (Publisher)
keywords: vbapb10.chm3735568
f1_keywords:
- vbapb10.chm3735568
ms.prod: publisher
api_name:
- Publisher.TextEffectFormat.ToggleVerticalText
ms.assetid: 627ddbcc-5951-70c6-4e54-de0e9a4bebec
ms.date: 06/08/2017
localization_priority: Normal
---


# TextEffectFormat.ToggleVerticalText Method (Publisher)

Switches the text flow in the specified WordArt from horizontal to vertical, or vice versa.


## Syntax

 _expression_. **ToggleVerticalText**

 _expression_ A variable that represents a  **TextEffectFormat** object.


## Remarks

Using the  **ToggleVerticalText** method swaps the values of the **[Left](Publisher.Shape.Left.md)** and **[Top](Publisher.Shape.Top.md)** properties of the **[Shape](Publisher.Shape.md)** object that represents the WordArt and leaves the  **[Width](Publisher.Shape.Width.md)** and **[Height](Publisher.Shape.Height.md)** properties unchanged.

The  **[Flip](Publisher.Shape.Flip.md)** method and  **[Rotation](Publisher.Shape.Rotation.md)** property of the  **[Shape](Publisher.Shape.md)** object and the  **[RotatedChars](Publisher.TextEffectFormat.RotatedChars.md)** property and  **ToggleVerticalText** method of the **[TextEffectFormat](Publisher.TextEffectFormat.md)** object all affect the character orientation and the direction of text flow in a **Shape** object that represents WordArt. You may have to experiment to find out how to combine the effects of these properties and methods to get the result you want.


## Example

This example adds WordArt that contains the text "Test" to the active publication, and switches from horizontal text flow (the default for the specified WordArt style,  **msoTextEffect1**) to vertical text flow.


```vb
Dim shpTextEffect As Shape 
 
Set shpTextEffect = ActiveDocument.Pages(1).Shapes.AddTextEffect _ 
 (PresetTextEffect:=msoTextEffect1, Text:="Test", _ 
 FontName:="Arial Black", FontSize:=36, _ 
 FontBold:=False, FontItalic:=False, Left:=100, Top:=100) 
 
shpTextEffect.TextEffect.ToggleVerticalText
```


