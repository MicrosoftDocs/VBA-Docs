---
title: Document.ResetFormFields method (Word)
keywords: vbawd10.chm158007671
f1_keywords:
- vbawd10.chm158007671
ms.prod: word
api_name:
- Word.Document.ResetFormFields
ms.assetid: 77354799-7ba7-a4e1-5379-c7664c8820b0
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.ResetFormFields method (Word)

Clears all form fields in a document, preparing the form to be filled in again.


## Syntax

_expression_. `ResetFormFields`

_expression_ Required. A variable that represents a **[Document](Word.Document.md)** object.


## Remarks

Use the  **ResetFormFields** method to clear fields when a document's fields are not locked. To clear fields when a document's fields are locked, use the **Protect** method.


## Example

This example clears the fields in the active document. This example assumes that the active document contains form fields.


```vb
Sub ClearFormFields() 
 ActiveDocument.ResetFormFields 
End Sub
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]