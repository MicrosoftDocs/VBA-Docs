---
title: Font.ColorIndex property (Word)
keywords: vbawd10.chm156369033
f1_keywords:
- vbawd10.chm156369033
ms.prod: word
api_name:
- Word.Font.ColorIndex
ms.assetid: c5011017-bf7a-5d89-0f20-f000d3ffd0ea
ms.date: 06/08/2017
localization_priority: Normal
---


# Font.ColorIndex property (Word)

Returns or sets a  **WdColorIndex** constant that represents the color for the specified font. Read/write .


## Syntax

_expression_.**ColorIndex**

_expression_ Required. A variable that represents a **[Font](Word.Font.md)** object.


## Remarks

The **wdByAuthor** constant is not a valid color for fonts.


## Example

This example changes the color of the text in the first paragraph in the active document.


```vb
ActiveDocument.Paragraphs(1).Range.Font.ColorIndex = wdGreen
```

This example formats the selected text to appear in red.




```vb
Selection.Font.ColorIndex = wdRed
```


## See also


[Font Object](Word.Font.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
