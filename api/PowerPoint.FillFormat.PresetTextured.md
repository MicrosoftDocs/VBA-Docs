---
title: FillFormat.PresetTextured method (PowerPoint)
keywords: vbapp10.chm552006
f1_keywords:
- vbapp10.chm552006
ms.prod: powerpoint
api_name:
- PowerPoint.FillFormat.PresetTextured
ms.assetid: a025a1d3-a2db-e219-7080-1a29c2fd3f21
ms.date: 06/08/2017
localization_priority: Normal
---


# FillFormat.PresetTextured method (PowerPoint)

Sets the specified fill to a preset texture.


## Syntax

_expression_.**PresetTextured** (_PresetTexture_)

_expression_ A variable that represents a **[FillFormat](powerpoint.fillformat.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _PresetTexture_|Required|**[MsoPresetTexture](Office.MsoPresetTexture.md)**|The preset texture.|



## Example

This example adds a rectangle with a green-marble textured fill to _myDocument_.


```vb
Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes.AddShape(msoShapeCan, 90, 90, 40, 80) _
    .Fill.PresetTextured msoTextureGreenMarble
```


## See also


[FillFormat Object](PowerPoint.FillFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]