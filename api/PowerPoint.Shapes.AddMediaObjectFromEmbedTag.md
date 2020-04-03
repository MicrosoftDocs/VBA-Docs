---
title: Shapes.AddMediaObjectFromEmbedTag method (PowerPoint)
keywords: vbapp10.chm543033
f1_keywords:
- vbapp10.chm543033
ms.prod: powerpoint
api_name:
- PowerPoint.Shapes.AddMediaObjectFromEmbedTag
ms.assetid: c463e7e2-8bac-8762-fec8-e1e84847907b
ms.date: 06/08/2017
localization_priority: Normal
---


# Shapes.AddMediaObjectFromEmbedTag method (PowerPoint)

Adds a media object from an embedded tag to a **Shapes** object.


## Syntax

_expression_. `AddMediaObjectFromEmbedTag`( `_EmbedTag_`, `_Left_`, `_Top_`, `_Width_`, `_Height_` )

_expression_ A variable that represents a **[Shapes](PowerPoint.Shapes.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _EmbedTag_|Required|**String**|The embed tag.|
| _Left_|Optional|**Single**|The distance, in [points](../language/glossary/vbe-glossary.md#point), from the left edge of the slide to the left edge of the media object.|
| _Top_|Optional|**Single**|The distance, in [points](../language/glossary/vbe-glossary.md#point), from the top edge of the slide to the top edge of the media object.|
| _Width_|Optional|**Single**|The width, in [points](../language/glossary/vbe-glossary.md#point), of the media object.|
| _Height_|Optional|**Single**|The height, in [points](../language/glossary/vbe-glossary.md#point), of the media object.|

## Return value

Shape


## See also


[Shapes Object](PowerPoint.Shapes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]