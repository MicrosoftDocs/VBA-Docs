---
title: Document.EndReview method (Word)
keywords: vbawd10.chm158007652
f1_keywords:
- vbawd10.chm158007652
ms.prod: word
api_name:
- Word.Document.EndReview
ms.assetid: bf53cefd-98e9-7e1a-016e-abd0c16e8bcd
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.EndReview method (Word)

Terminates a review of a file that has been sent for review using the **[SendForReview](Word.Document.SendForReview.md)** method or that has been automatically placed in a review cycle by sending a document to another user in an email message.


## Syntax

_expression_. `EndReview`

_expression_ Required. A variable that represents a **[Document](Word.Document.md)** object.


## Remarks

When executed, the **EndReview** method displays a message asking the user whether to end the review.


## Example

This example terminates the review of the active document. This example assumes the active document part of a review cycle.


```vb
Sub EndDocRev() 
 ActiveDocument.EndReview 
End Sub
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]