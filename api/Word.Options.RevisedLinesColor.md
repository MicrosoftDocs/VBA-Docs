---
title: Options.RevisedLinesColor property (Word)
keywords: vbawd10.chm162988094
f1_keywords:
- vbawd10.chm162988094
ms.prod: word
api_name:
- Word.Options.RevisedLinesColor
ms.assetid: bc8cd36f-49ac-119a-4f9f-f2e9b20f9bd6
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.RevisedLinesColor property (Word)

Returns or sets the color of changed lines in a document with tracked changes. Read/write  **WdColorIndex**.


## Syntax

_expression_. `RevisedLinesColor`

_expression_ Required. A variable that represents an **[Options](Word.Options.md)** object.


## Example

This example sets the color of changed lines to pink.


```vb
Options.RevisedLinesColor = wdPink
```

This example returns the current status of the  **Color** option under **Changed lines** on the **Track Changes** tab in the **Options** dialog box (**Tools** menu).




```vb
temp = Options.RevisedLinesColor
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]