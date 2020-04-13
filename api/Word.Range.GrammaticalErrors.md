---
title: Range.GrammaticalErrors property (Word)
keywords: vbawd10.chm157155643
f1_keywords:
- vbawd10.chm157155643
ms.prod: word
api_name:
- Word.Range.GrammaticalErrors
ms.assetid: 2535ba4d-1c5c-3dc2-2ddc-14c8a5625f41
ms.date: 06/08/2017
localization_priority: Normal
---


# Range.GrammaticalErrors property (Word)

Returns a  **[ProofreadingErrors](Word.proofreadingerrors.md)** collection that represents the sentences that failed the grammar check on the specified document or range. Read-only.


## Syntax

_expression_. `GrammaticalErrors`

_expression_ A variable that represents a **[Range](Word.Range.md)** object.


## Remarks

There can be more than one error per sentence. If there are no grammatical errors, the **Count** property for the **ProofreadingErrors** object returned by the **GrammaticalErrors** property returns 0 (zero).

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example checks the third paragraph in the active document for grammatical errors and displays each sentence that contains one or more errors.


```vb
Set myErrors = ActiveDocument.Paragraphs(3).Range.GrammaticalErrors 
For Each myerr In myErrors 
 MsgBox myerr.Text 
Next myerr
```


## See also


[Range Object](Word.Range.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]