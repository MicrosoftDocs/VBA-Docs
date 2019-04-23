---
title: Conflicts object (Word)
ms.prod: word
api_name:
- Word.Conflicts
ms.assetid: 476e8f6d-c93e-b372-2fa7-1c9a4a84a182
ms.date: 06/08/2017
localization_priority: Normal
---


# Conflicts object (Word)

 A collection of[Conflict](Word.Conflict.md) objects that represents the conflicts in a document. The type of a **Conflict** object is specified by the [WdRevisionType](Word.WdRevisionType.md) enumeration.


## Remarks

Use the [Conflicts](Word.CoAuthoring.Conflicts.md) property to return the **Conflicts** collection for a document. Use Conflicts (_index_), where _index_ is the conflict index number, to return a single **Conflict** object.


## Example

The following code example accepts the first conflict in the active document.


```vb
ActiveDocument.CoAuthoring.Conflicts(1).Accept 

```


## See also



[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]