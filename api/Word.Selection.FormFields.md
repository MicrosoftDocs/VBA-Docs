---
title: Selection.FormFields property (Word)
keywords: vbawd10.chm158662721
f1_keywords:
- vbawd10.chm158662721
ms.prod: word
api_name:
- Word.Selection.FormFields
ms.assetid: d6d5259b-9971-929f-16f7-ca2b2d585c77
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.FormFields property (Word)

Returns a  **[FormFields](Word.formfields.md)** collection that represents all the form fields in the selection. Read-only.


## Syntax

_expression_. `FormFields`

_expression_ A variable that represents a **[Selection](Word.Selection.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example displays the name of the first form field in the selection.


```vb
If Selection.FormFields.Count > 0 Then 
 MsgBox Selection.FormFields(1).Name 
End If
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]