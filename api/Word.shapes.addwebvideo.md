---
title: Shapes.AddWebVideo method (Word)
keywords: vbawd10.chm161415272
f1_keywords:
- vbawd10.chm161415272
ms.prod: word
ms.assetid: 9bdd1bc2-0d04-ca0c-eba2-4080843cf614
ms.date: 06/08/2017
localization_priority: Normal
---


# Shapes.AddWebVideo method (Word)

Adds a new web video to the document.


## Syntax

_expression_.**AddWebVideo** (_EmbedCode_, _VideoWidth_, _VideoHeight_, _PosterFrameImage_, _Url_, _Left_, _Top_, _Width_, _Height_, _Anchor_)

_expression_ A variable that represents a **[Shapes](Word.shapes.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _EmbedCode_|Required|**String**|The HTML code to embed.|
| _VideoWidth_|Required|**Variant**|An integer that represents the width of the web video in pixels.|
| _VideoHeight_|Required|**Variant**|An integer that represents the height of the web video in pixels.|
| _PosterFrameImage_|Optional|**Variant**|A string that points to the file to use as the poster frame for the web video.|
| _Url_|Optional|**Variant**|A string that contains the URL to the web video.|
| _Left_|Optional|**Variant**|The position, measured in points, of the left edge of the poster frame from the edge of the document.|
| _Top_|Optional|**Variant**|The position, measured in points, of the top edge of the poster frame from the edge of the document.|
| _Width_|Optional|**Variant**|The width, measured in points, of the poster frame in the document.|
| _Height_|Optional|**Variant**|The height, measured in points, of the poster frame in the document.|
| _Anchor_|Optional|**Variant**|A **[Range](Word.Range.md)** object that represents the text to which the web video is bound. If _Anchor_ is specified, the anchor is positioned at the beginning of the first paragraph in the anchoring range. If this argument is omitted, the anchoring range is selected automatically and the video is positioned relative to the top and left edges of the page.|

## Return value

**SHAPE**



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]