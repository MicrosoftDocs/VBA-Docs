---
title: InlineShapes.AddHorizontalLineStandard method (Word)
keywords: vbawd10.chm162070633
f1_keywords:
- vbawd10.chm162070633
ms.prod: word
api_name:
- Word.InlineShapes.AddHorizontalLineStandard
ms.assetid: de9d4613-4e64-9df8-aa9a-890335eb648d
ms.date: 06/08/2017
localization_priority: Normal
---


# InlineShapes.AddHorizontalLineStandard method (Word)

Adds a horizontal line to the current document.


## Syntax

_expression_. `AddHorizontalLineStandard`( `_Range_` )

_expression_ Required. A variable that represents an '[InlineShapes](Word.inlineshapes.md)' collection.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Range_|Optional| **Variant**|The range above which Microsoft Word places the horizontal line. If this argument is omitted, Word places the horizontal line above the current selection.|

## Remarks

To add a horizontal line based on an existing image file, use the  **[AddHorizontalLine](Word.InlineShapes.AddHorizontalLine.md)** method.


## Example

This example adds a horizontal line above the fifth paragraph in the active document.


```vb
ActiveDocument.Paragraphs(5).Range _ 
 .InlineShapes.AddHorizontalLineStandard
```


## See also


[InlineShapes Collection Object](Word.inlineshapes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]