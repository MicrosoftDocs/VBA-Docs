---
title: InlineShapes.AddWebVideo method (Word)
keywords: vbawd10.chm162070637
f1_keywords:
- vbawd10.chm162070637
ms.prod: word
ms.assetid: b91c763e-9865-5591-7c90-6eafe1a1848a
ms.date: 06/08/2017
localization_priority: Normal
---


# InlineShapes.AddWebVideo method (Word)

Adds a new web video to the document.


## Syntax

_expression_. `AddWebVideo`_(EmbedCode,_ _VideoWidth,_ _VideoHeight,_ _PosterFrameImage,_ _Url,_ _Range)_

 _expression_ A variable that represents a 'InlineShapes' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|||||
| _EmbedCode_|Required|**String**|The embed code for the video.|
| _VideoWidth_|Required|**Variant**|An integer that represents the width of the web video in pixels.|
| _VideoHeight_|Required|**Variant**|An integer that represents the height of the web video in pixels.|
| _PosterFrameImage_|Optional|**Variant**|A string that points to the file to use as the poster frame for the web video.|
| _Url_|Optional|**Variant**|The URL to the video.|
| _Range_|Optional|**Variant**|The range at which to insert the web video. If  _Range_ is omitted, the current selection is used.|

## Return value

 **INLINESHAPE**


## See also


[InlineShapes Collection](Word.inlineshapes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]