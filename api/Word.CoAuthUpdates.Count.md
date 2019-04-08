---
title: CoAuthUpdates.Count property (Word)
keywords: vbawd10.chm217841665
f1_keywords:
- vbawd10.chm217841665
ms.prod: word
api_name:
- Word.CoAuthUpdates.Count
ms.assetid: a0918742-9fbf-2a57-8efd-1487dd56d451
ms.date: 06/08/2017
localization_priority: Normal
---


# CoAuthUpdates.Count property (Word)

Returns the number of items in the [CoAuthUpdates](overview/Word.md) collection. Read-only.


## Syntax

_expression_.**Count**

 _expression_ An expression that returns a 'CoAuthUpdates' object.


## Example

The following code example displays the number of content updates that were merged into the active document at the last explicit save.


```vb
MsgBox "The active document contains " & _ 
    ActiveDocument.CoAuthoring.Updates.Count & " update(s)."
```


## See also


[CoAuthUpdates Object](overview/Word.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]