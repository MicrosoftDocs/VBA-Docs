---
title: Shapes.AddWordArt method (Publisher)
keywords: vbapb10.chm2162761
f1_keywords:
- vbapb10.chm2162761
ms.prod: publisher
api_name:
- Publisher.Shapes.AddWordArt
ms.assetid: 8ff83baa-5d88-5f80-3a69-5f712ba5e583
ms.date: 06/14/2019
localization_priority: Normal
---


# Shapes.AddWordArt method (Publisher)

Returns a **[Shape](Publisher.Shape.md)** object that represents the WordArt to be added to the publication.


## Syntax

_expression_.**AddWordArt** (_PresetWordArt_, _Text_, _FontName_, _FontSize_, _FontBold_, _FontItalic_, _Left_, _Top_)

_expression_ A variable that represents a **[Shapes](Publisher.Shapes.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _PresetWordArt_|Required| **[PbPresetWordArt](publisher.pbpresetwordart.md)**|The type of preset WordArt to add.|
| _Text_ |Required| **String**|The text of the WordArt.|
| _FontName_ |Required| **String**|The name of the font to be used in the WordArt.|
| _FontSize_ |Required| **Variant**|The size of the font to be used in the WordArt.|
| _FontBold_ |Required| **[MsoTriState](office.msotristate.md)**|Whether the WordArt text should be bold. See Remarks for possible values.|
| _FontItalic_ |Required| **MsoTriState**|Whether the WordArt text should be italic. See Remarks for possible values.|
| _Left_ |Required| **Variant**|The horizontal position of the WordArt.|
| _Top_ |Required| **Variant**|The vertical position of the WordArt.|

## Return value

Shape


## Remarks

The _FontBold_ parameter value can be one of the following **MsoTriState** constants declared in the Microsoft Office type library.

|Constant|Description|
|:-----|:-----|
| **msoFalse**|None of the characters in the WordArt are formatted as bold.|
| **msoTriStateMixed**|A return value indicating that the WordArt contains some text formatted as bold and some text not formatted as bold.|
| **msoTriStateToggle**|A set value that switches between **msoTrue** and **msoFalse**.|
| **msoTrue**|All the characters in the WordArt are formatted as bold.|

<br/>

The _FontItalic_ parameter value can be one of the following **MsoTriState** constants.

|Constant|Description|
|:-----|:-----|
| **msoFalse**|None of the characters in the WordArt are formatted as italic.|
| **msoTriStateMixed**|A return value indicating that the WordArt contains some text formatted as italic and some text not formatted as italic.|
| **msoTriStateToggle**|A set value that switches between **msoTrue** and **msoFalse**.|
| **msoTrue**|All the characters in the WordArt are formatted as italic.|


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]