---
title: Options.CursorMovement property (Word)
keywords: vbawd10.chm162988435
f1_keywords:
- vbawd10.chm162988435
ms.prod: word
api_name:
- Word.Options.CursorMovement
ms.assetid: f73f8a6e-4a66-e3f8-7197-42d5c1f73bcf
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.CursorMovement property (Word)

Returns or sets how the insertion point progresses within bidirectional text. Read/write  **WdCursorMovement**.


## Syntax

_expression_. `CursorMovement`

_expression_ Required. A variable that represents an **[Options](Word.Options.md)** object.


## Example

This example sets the insertion point to progress to the next visually adjacent character as it moves through bidirectional text.


```vb
Options.CursorMovement = wdCursorMovementVisual
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]