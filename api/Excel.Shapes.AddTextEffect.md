---
title: Shapes.AddTextEffect method (Excel)
keywords: vbaxl10.chm638085
f1_keywords:
- vbaxl10.chm638085
ms.prod: excel
api_name:
- Excel.Shapes.AddTextEffect
ms.assetid: ace2bd71-455d-d187-7fb0-77eed879ff95
ms.date: 05/15/2019
localization_priority: Normal
---


# Shapes.AddTextEffect method (Excel)

Creates a WordArt object. Returns a **[Shape](Excel.Shape.md)** object that represents the new WordArt object.


## Syntax

_expression_.**AddTextEffect** (_PresetTextEffect_, _Text_, _FontName_, _FontSize_, _FontBold_, _FontItalic_, _Left_, _Top_)

_expression_ A variable that represents a **[Shapes](Excel.Shapes.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _PresetTextEffect_|Required| **[MsoPresetTextEffect](Office.MsoPresetTextEffect.md)**|The preset text effect.|
| _Text_|Required| **String**|The text in the WordArt.|
| _FontName_|Required| **String**|The name of the font used in the WordArt.|
| _FontSize_|Required| **Single**|The size (in [points](../language/glossary/vbe-glossary.md#point)) of the font used in the WordArt.|
| _FontBold_|Required| **[MsoTriState](Office.MsoTriState.md)**|The font used in the WordArt to bold.|
| _FontItalic_|Required| **MsoTriState**|The font used in the WordArt to italic.|
| _Left_|Required| **Single**|The position (in points) of the upper-left corner of the WordArt's bounding box relative to the upper-left corner of the document.|
| _Top_|Required| **Single**|The position (in points) of the upper-left corner of the WordArt's bounding box relative to the top of the document.|

## Return value

**Shape**


## Remarks

When you add WordArt to a document, the height and width of the WordArt are automatically set based on the size and amount of text that you specify.


## Example

This example adds WordArt that contains the text Test to _myDocument_.

```vb
Set myDocument = Worksheets(1) 
Set newWordArt = myDocument.Shapes.AddTextEffect( _ 
    PresetTextEffect:=msoTextEffect1, Text:="Test", _ 
    FontName:="Arial Black", FontSize:=36, _ 
    FontBold:=msoFalse, FontItalic:=msoFalse, Left:=10, _ 
    Top:=10)
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]