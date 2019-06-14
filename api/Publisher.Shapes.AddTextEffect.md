---
title: Shapes.AddTextEffect method (Publisher)
keywords: vbapb10.chm2162721
f1_keywords:
- vbapb10.chm2162721
ms.prod: publisher
api_name:
- Publisher.Shapes.AddTextEffect
ms.assetid: 21af82f1-d507-3c16-72df-bde1b5e00717
ms.date: 06/14/2019
localization_priority: Normal
---


# Shapes.AddTextEffect method (Publisher)

Adds a new **[Shape](Publisher.Shape.md)** object representing a WordArt object to the specified **Shapes** collection.


## Syntax

_expression_.**AddTextEffect** (_PresetTextEffect_, _Text_, _FontName_, _FontSize_, _FontBold_, _FontItalic_, _Left_, _Top_)

_expression_ A variable that represents a **[Shapes](Publisher.Shapes.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_PresetTextEffect_|Required| **[MsoPresetTextEffect](Office.MsoPresetTextEffect.md)**|The preset text effect to use. The values of the **MsoPresetTextEffect** constants correspond to the formats listed in the **WordArt Gallery** dialog box (numbered from left to right and from top to bottom). Can be one of the **MsoPresetTextEffect** constants declared in the Microsoft Office type library. The **msoTextEffectMixed** constant is not supported.|
| _Text_ |Required| **String**|The text to use for the WordArt object.|
| _FontName_ |Required| **String**|The name of the font to use for the WordArt object.|
| _FontSize_ |Required| **Variant**|The font size to use for the WordArt object. Numeric values are evaluated in [points](../language/glossary/vbe-glossary.md#point); strings can be in any units supported by Microsoft Publisher (for example, "2.5 in").|
| _FontBold_ |Required| **[MsoTriState](office.msotristate.md)**|Determines whether to format the WordArt text as bold.|
| _FontItalic_ |Required| **MsoTriState**|Determines whether to format the WordArt text as italic.|
| _Left_ |Required| **Variant**|The position of the left edge of the shape representing the WordArt object.|
| _Top_ |Required| **Variant**|The position of the top edge of the shape representing the WordArt object.|

## Return value

Shape


## Remarks

For the _Left_ and _Top_ parameters, numeric values are evaluated in [points](../language/glossary/vbe-glossary.md#point); strings can be in any units supported by Publisher (for example, "2.5 in").

The height and width of the WordArt object is determined by its text and formatting.

Use the **[Shape.TextEffect](Publisher.Shape.TextEffect.md)** property to return a **[TextEffectFormat](Publisher.TextEffectFormat.md)** object whose properties can be used to edit an existing WordArt object.

The _FontBold_ parameter can be one of the **MsoTriState** constants declared in the Microsoft Office type library and shown in the following table.

|Constant|Description|
|:-----|:-----|
| **msoFalse**|Do not format the WordArt text as bold.|
| **msoTrue** |Format the WordArt text as bold.|

The _FontItalic_ parameter can be one of the **MsoTriState** constants shown in the following table.

|Constant|Description|
|:-----|:-----|
| **msoFalse**|Do not format the WordArt text as italic.|
| **msoTrue** |Format the WordArt text as italic.|

## Example

The following example adds a WordArt object to the first page of the active publication.

```vb
Dim shpWordArt As Shape 
 
Set shpWordArt = ActiveDocument.Pages(1).Shapes.AddTextEffect _ 
 (PresetTextEffect:=msoTextEffect7, Text:="Annual Report", _ 
 FontName:="Arial Black", FontSize:=24, _ 
 FontBold:=msoFalse, FontItalic:=msoFalse, _ 
 Left:=144, Top:=72) 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]