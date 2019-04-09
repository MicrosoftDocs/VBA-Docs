---
title: InlineShapes.AddPictureBullet method (Word)
keywords: vbawd10.chm162070634
f1_keywords:
- vbawd10.chm162070634
ms.prod: word
api_name:
- Word.InlineShapes.AddPictureBullet
ms.assetid: 39e6ea87-eddf-5c08-07bf-52bd13de1117
ms.date: 06/08/2017
localization_priority: Normal
---


# InlineShapes.AddPictureBullet method (Word)

Adds a picture bullet based on an image file to the current document. Returns an  **[InlineShape](Word.InlineShape.md)** object.


## Syntax

_expression_. `AddPictureBullet`( `_FileName_` , `_Range_` )

_expression_ Required. A variable that represents an '[InlineShapes](Word.inlineshapes.md)' collection.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **String**|The file name of the image you want to use for the picture bullet.|
| _Range_|Optional| **Variant**|The range to which Microsoft Word adds the picture bullet. Word adds the picture bullet to each paragraph in the range. If this argument is omitted, Word adds the picture bullet to each paragraph in the current selection.|

## Return value

InlineShape


## Example

This example adds a picture bullet to each paragraph in the selected text using a file called "RedBullet.gif."


```vb
Selection.InlineShapes.AddPictureBullet _ 
 "C:\Art files\RedBullet.gif"
```


## See also


[InlineShapes Collection Object](Word.inlineshapes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]