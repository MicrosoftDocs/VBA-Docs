---
title: FillFormat.PresetTextured method (Publisher)
keywords: vbapb10.chm2359316
f1_keywords:
- vbapb10.chm2359316
ms.prod: publisher
api_name:
- Publisher.FillFormat.PresetTextured
ms.assetid: 971eac34-4e29-c898-93c8-9e71bd92238d
ms.date: 06/07/2019
localization_priority: Normal
---


# FillFormat.PresetTextured method (Publisher)

Sets the specified fill to a preset texture.


## Syntax

_expression_.**PresetTextured** (_PresetTexture_)

_expression_ A variable that represents a **[FillFormat](publisher.fillformat.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_PresetTexture_ |Required| **[MsoPresetTexture](Office.MsoPresetTexture.md)** |The preset texture. Can be one of the **MsoPresetTexture** constants declared in the Microsoft Office type library.|



## Example

This example adds a rectangle with a green-marble textured fill to the active publication.

```vb
ActiveDocument.Pages(1).Shapes _ 
 .AddShape(Type:=msoShapeCan, _ 
 Left:=90, Top:=90, Width:=40, Height:=80) _ 
 .Fill.PresetTextured _ 
 PresetTexture:=msoTextureGreenMarble 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]