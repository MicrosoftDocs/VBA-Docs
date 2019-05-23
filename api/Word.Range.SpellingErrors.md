---
title: Range.SpellingErrors property (Word)
keywords: vbawd10.chm157155644
f1_keywords:
- vbawd10.chm157155644
ms.prod: word
api_name:
- Word.Range.SpellingErrors
ms.assetid: 4b35a13d-2a5f-e9cd-0667-58aae00a48f1
ms.date: 06/08/2017
localization_priority: Normal
---


# Range.SpellingErrors property (Word)

Returns a  **ProofreadingErrors** collection that represents the words identified as spelling errors in the specified range. Read-only.


## Syntax

_expression_. `SpellingErrors`

_expression_ A variable that represents a **[Range](Word.Range.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning a Single Object from a Collection](../word/Concepts/Miscellaneous/returning-a-single-object-from-a-collection.md).


## Example

This example checks the specified range for spelling errors and displays each error found.


```vb
Set myErrors = ActiveDocument.Paragraphs(3).Range.SpellingErrors 
If myErrors.Count = 0 Then 
 Msgbox "No spelling errors found." 
Else 
 For Each myErr in myErrors 
 Msgbox myErr.Text 
 Next 
End If
```


## See also


[Range Object](Word.Range.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]