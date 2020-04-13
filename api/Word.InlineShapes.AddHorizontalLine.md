---
title: InlineShapes.AddHorizontalLine method (Word)
keywords: vbawd10.chm162070632
f1_keywords:
- vbawd10.chm162070632
ms.prod: word
api_name:
- Word.InlineShapes.AddHorizontalLine
ms.assetid: d35591f3-7a42-e4e1-0532-ef1b3b44803a
ms.date: 06/08/2017
localization_priority: Normal
---


# InlineShapes.AddHorizontalLine method (Word)

Adds a horizontal line based on an image file to the current document.


## Syntax

_expression_. `AddHorizontalLine`( `_FileName_` , `_Range_` )

_expression_ Required. A variable that represents an '[InlineShapes](Word.inlineshapes.md)' collection.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **String**|The file name of the image you want to use for the horizontal line.|
| _Range_|Optional| **Variant**|The range above which Microsoft Word places the horizontal line. If this argument is omitted, Word places the horizontal line above the current selection.|

## Remarks

To add a horizontal line that isn't based on an existing image file, use the **[AddHorizontalLineStandard](Word.InlineShapes.AddHorizontalLineStandard.md)** method.


## Example

This example adds a horizontal line above the current selection in the active document using a file called "ArtsyRule.gif."


```vb
Selection.InlineShapes.AddHorizontalLine _ 
 "C:\Art files\ArtsyRule.gif"
```


## See also


[InlineShapes Collection Object](Word.inlineshapes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]