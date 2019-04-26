---
title: FillFormat.PresetTextured method (Word)
keywords: vbawd10.chm164102158
f1_keywords:
- vbawd10.chm164102158
ms.prod: word
api_name:
- Word.FillFormat.PresetTextured
ms.assetid: 9a4aac9d-6349-7947-bc4a-1b0bb64a848b
ms.date: 06/08/2017
localization_priority: Normal
---


# FillFormat.PresetTextured method (Word)

Sets the specified fill to a preset texture.


## Syntax

_expression_.**PresetTextured** (_PresetTexture_)

_expression_ Required. A variable that represents a **[FillFormat](word.fillformat.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _PresetTexture_|Required| **[MsoPresetTexture](Office.MsoPresetTexture.md)**|The preset texture.|

## Example

This example adds a rectangle with a green-marble textured fill to the active document.


```vb
ActiveDocument.Shapes.AddShape(msoShapeCan, 90, 90, 40, 80) _ 
 .Fill.PresetTextured msoTextureGreenMarble
```


## See also


[FillFormat Object](Word.FillFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]