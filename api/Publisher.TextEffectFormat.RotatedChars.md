---
title: TextEffectFormat.RotatedChars property (Publisher)
keywords: vbapb10.chm3735817
f1_keywords:
- vbapb10.chm3735817
ms.prod: publisher
api_name:
- Publisher.TextEffectFormat.RotatedChars
ms.assetid: 47566497-7b78-65dc-48d9-26b2e4245d31
ms.date: 06/15/2019
localization_priority: Normal
---


# TextEffectFormat.RotatedChars property (Publisher)

Returns **msoTrue** if characters in the specified WordArt are rotated 90 degrees relative to the WordArt's bounding shape. Returns **msoFalse** if characters in the specified WordArt retain their original orientation relative to the bounding shape. Read/write.


## Syntax

_expression_.**RotatedChars**

_expression_ A variable that represents a **[TextEffectFormat](Publisher.TextEffectFormat.md)** object.


## Return value

**[MsoTriState](Office.MsoTriState.md)**


## Remarks

If the WordArt has horizontal text, setting the **RotatedChars** property to **True** rotates the characters 90 degrees counterclockwise. 

If the WordArt has vertical text, setting the **RotatedChars** property to **False** rotates the characters 90 degrees clockwise. 

Use the **[ToggleVerticalText](Publisher.TextEffectFormat.ToggleVerticalText.md)** method to switch between horizontal and vertical text flow.

The **[Flip](Publisher.Shape.Flip.md)** method and **[Rotation](Publisher.Shape.Rotation.md)** property of the **Shape** object and the **RotatedChars** property and **ToggleVerticalText** method all affect the character orientation and direction of text flow in a **Shape** object that represents WordArt. You may have to experiment to find out how to combine the effects of these properties and methods to get the result you want.


## Example

This example adds WordArt that contains the text Test to the active publication and rotates the characters 90 degrees counterclockwise.

```vb
Sub CreateFormatWordArt() 
 With ActiveDocument.Pages(1).Shapes _ 
 .AddTextEffect(PresetTextEffect:=msoTextEffect1, _ 
 Text:="Test", FontName:="Arial Black", FontSize:=36, _ 
 FontBold:=False, FontItalic:=False, Left:=10, Top:=10) 
 .TextEffect.RotatedChars = msoTrue 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]