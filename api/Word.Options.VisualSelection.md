---
title: Options.VisualSelection property (Word)
keywords: vbawd10.chm162988436
f1_keywords:
- vbawd10.chm162988436
ms.prod: word
api_name:
- Word.Options.VisualSelection
ms.assetid: d3947a4c-0495-6211-7646-3b202855d35a
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.VisualSelection property (Word)

Returns or sets the selection behavior based on visual cursor movement in a right-to-left language document. Read/write  **WdVisualSelection**.


## Syntax

_expression_. `VisualSelection`

_expression_ Required. A variable that represents an **[Options](Word.Options.md)** object.


## Remarks

The **CursorMovement** property must be set to **wdCursorMovementVisual** to use this property.


## Example

This example sets the selection behavior so that the selection wraps from line to line.


```vb
If Options.CursorMovement = wdCursorMovementVisual Then _ 
 Options.VisualSelection = wdVisualSelectionContinuous
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]