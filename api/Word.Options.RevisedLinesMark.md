---
title: Options.RevisedLinesMark property (Word)
keywords: vbawd10.chm162988091
f1_keywords:
- vbawd10.chm162988091
ms.prod: word
api_name:
- Word.Options.RevisedLinesMark
ms.assetid: ecc358f2-4bf6-7546-5400-938a3dae6b77
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.RevisedLinesMark property (Word)

Returns or sets the placement of changed lines in a document with tracked changes. Read/write  **WdRevisedLinesMark**.


## Syntax

_expression_. `RevisedLinesMark`

_expression_ Required. A variable that represents an **[Options](Word.Options.md)** object.


## Example

This example sets changed lines to appear in the left margin of every page.


```vb
Options.RevisedLinesMark = wdRevisedLinesMarkLeftBorder
```

This example returns the current status of the  **Mark** option under **Changed lines** on the **Track Changes** tab in the **Options** dialog box (**Tools** menu).




```vb
temp = Options.RevisedLinesMark
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]