---
title: Selection.Hyperlinks property (Word)
keywords: vbawd10.chm158662812
f1_keywords:
- vbawd10.chm158662812
ms.prod: word
api_name:
- Word.Selection.Hyperlinks
ms.assetid: c90c3779-cbb9-4174-3002-850750b4bb41
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.Hyperlinks property (Word)

Returns a  **[Hyperlinks](Word.hyperlinks.md)** collection that represents all the hyperlinks in the specified selection. Read-only.


## Syntax

_expression_.**Hyperlinks**

_expression_ A variable that represents a **[Selection](Word.Selection.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example jumps to the address of the first hyperlink in the selection.


```vb
If Selection.Hyperlinks.Count >= 1 Then 
 Selection.Hyperlinks(1).Follow 
End If
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]