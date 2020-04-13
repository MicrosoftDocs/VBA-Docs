---
title: Document.GrammaticalErrors property (Word)
keywords: vbawd10.chm158007393
f1_keywords:
- vbawd10.chm158007393
ms.prod: word
api_name:
- Word.Document.GrammaticalErrors
ms.assetid: 24e708e3-6417-f105-43d3-9be8e450f189
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.GrammaticalErrors property (Word)

Returns a  **[ProofreadingErrors](Word.proofreadingerrors.md)** collection that represents the sentences that failed the grammar check in the specified document. Read-only.


## Syntax

_expression_. `GrammaticalErrors`

_expression_ A variable that represents a **[Document](Word.Document.md)** object.


## Remarks

There can be more than one error per sentence. If there are no grammatical errors, the **Count** property for the **[ProofreadingErrors](Word.proofreadingerrors.md)** collection returned by the **GrammaticalErrors** property returns 0 (zero).

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example checks the active document for grammatical errors. If any errors are found, a new spelling and grammar check is started.


```vb
If ActiveDocument.GrammaticalErrors.Count = 0 Then 
 Msgbox "There are no grammatical errors." 
Else 
 ActiveDocument.CheckGrammar 
End If
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]