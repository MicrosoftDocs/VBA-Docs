---
title: Document.SpellingErrors property (Word)
keywords: vbawd10.chm158007394
f1_keywords:
- vbawd10.chm158007394
ms.prod: word
api_name:
- Word.Document.SpellingErrors
ms.assetid: c8a987a1-3705-ea0a-103a-99b2f17f5c6b
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.SpellingErrors property (Word)

Returns a  **[ProofreadingErrors](Word.proofreadingerrors.md)** collection that represents the words identified as spelling errors in the specified document or range. Read-only.


## Syntax

_expression_. `SpellingErrors`

_expression_ A variable that represents a **[Document](Word.Document.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example checks the active document for spelling errors and displays the number of errors found.


```vb
myErr = ActiveDocument.SpellingErrors.Count 
If myErr = 0 Then 
 Msgbox "No spelling errors found." 
Else 
 Msgbox myErr & " spelling errors found." 
End If
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]