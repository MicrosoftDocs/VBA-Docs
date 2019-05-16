---
title: TextEffectFormat.ToggleVerticalText method (Excel)
keywords: vbaxl10.chm118020
f1_keywords:
- vbaxl10.chm118020
ms.prod: excel
api_name:
- Excel.TextEffectFormat.ToggleVerticalText
ms.assetid: 9b4312b8-1642-9a49-6395-b49b129f44f2
ms.date: 05/17/2019
localization_priority: Normal
---


# TextEffectFormat.ToggleVerticalText method (Excel)

Switches the text flow in the specified WordArt from horizontal to vertical, or vice versa.


## Syntax

_expression_.**ToggleVerticalText**

_expression_ A variable that represents a **[TextEffectFormat](Excel.TextEffectFormat.md)** object.


## Remarks

Using the **ToggleVerticalText** method swaps the values of the **[Width](Excel.Shape.Width.md)** and **[Height](Excel.Shape.Height.md)** properties of the **Shape** object that represents the WordArt, and leaves the **[Left](Excel.Shape.Left.md)** and **[Top](Excel.Shape.Top.md)** properties unchanged.

The **[Flip](Excel.Shape.Flip.md)** method and **[Rotation](Excel.Shape.Rotation.md)** property of the **Shape** object and the **[RotatedChars](Excel.TextEffectFormat.RotatedChars.md)** property and **ToggleVerticalText** method of the **TextEffectFormat** object all affect the character orientation and the direction of text flow in a **Shape** object that represents WordArt. You may have to experiment to find out how to combine the effects of these properties and methods to get the result you want.


## Example

This example adds WordArt that contains the text Test to _myDocument_ and switches from horizontal text flow (the default for the specified WordArt style, **msoTextEffect1**) to vertical text flow.

```vb
Set myDocument = Worksheets(1) 
Set newWordArt = myDocument.Shapes.AddTextEffect( _ 
 PresetTextEffect:=msoTextEffect1, Text:="Test", _ 
 FontName:="Arial Black", FontSize:=36, _ 
 FontBold:=False, FontItalic:=False, Left:=100, _ 
 Top:=100) 
newWordArt.TextEffect.ToggleVerticalText
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]