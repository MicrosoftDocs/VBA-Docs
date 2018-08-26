---
title: CoAuthUpdates.Count Property (Word)
keywords: vbawd10.chm217841665
f1_keywords:
- vbawd10.chm217841665
ms.prod: word
api_name:
- Word.CoAuthUpdates.Count
ms.assetid: a0918742-9fbf-2a57-8efd-1487dd56d451
ms.date: 06/08/2017
---


# CoAuthUpdates.Count Property (Word)

Returns the number of items in the [CoAuthUpdates](./overview/Word.md) collection. Read-only.


## Syntax

 _expression_. `Count`

 _expression_ An expression that returns a 'CoAuthUpdates' object.


## Example

The following code example displays the number of content updates that were merged into the active document at the last explicit save.


```vb
<<<<<<< HEAD
MsgBox "The active document contains " &; _ 
    ActiveDocument.CoAuthoring.Updates.Count &; " update(s)."
=======
MsgBox "The active document contains " & _ 
    ActiveDocument.CoAuthoring.Updates.Count & " update(s)."
>>>>>>> master
```


## See also


[CoAuthUpdates Object](./overview/Word.md)


