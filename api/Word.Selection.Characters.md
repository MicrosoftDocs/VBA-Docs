---
title: Selection.Characters property (Word)
keywords: vbawd10.chm158662709
f1_keywords:
- vbawd10.chm158662709
ms.prod: word
api_name:
- Word.Selection.Characters
ms.assetid: 605c0fc5-f5b9-6782-9fdd-54589040d243
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.Characters property (Word)

Returns a  **[Characters](Word.characters.md)** collection that represents the characters in a document, range, or selection. Read-only.


## Syntax

_expression_. `Characters`

_expression_ A variable that represents a **[Selection](Word.Selection.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example displays the first character in the selection. If nothing is selected, the character immediately after the insertion point is displayed.


```vb
char = Selection.Characters(1).Text 
MsgBox "The first character is... " & char
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]