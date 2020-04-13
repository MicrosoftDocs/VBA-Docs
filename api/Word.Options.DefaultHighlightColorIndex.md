---
title: Options.DefaultHighlightColorIndex property (Word)
keywords: vbawd10.chm162988306
f1_keywords:
- vbawd10.chm162988306
ms.prod: word
api_name:
- Word.Options.DefaultHighlightColorIndex
ms.assetid: 1171cc44-54c9-0a39-c90f-ebdebebdde26
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.DefaultHighlightColorIndex property (Word)

Returns or sets the color used to highlight text formatted with the **Highlight** button (**Formatting** toolbar). Read/write **WdColorIndex**.


## Syntax

_expression_. `DefaultHighlightColorIndex`

_expression_ Required. A variable that represents an **[Options](Word.Options.md)** object.


## Example

This example sets the default highlight color to bright green. The new color doesn't apply to any previously highlighted text.


```vb
Options.DefaultHighlightColorIndex = wdBrightGreen
```

This example returns the current default highlight color index.




```vb
Dim lngTemp As Long 
 
lngTemp = Options.DefaultHighlightColorIndex
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]