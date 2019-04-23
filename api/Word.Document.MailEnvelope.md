---
title: Document.MailEnvelope property (Word)
keywords: vbawd10.chm158007632
f1_keywords:
- vbawd10.chm158007632
ms.prod: word
api_name:
- Word.Document.MailEnvelope
ms.assetid: f37a52f5-ebfe-a9b9-056e-50f6adf4c1b4
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.MailEnvelope property (Word)

Returns an  **MsoEnvelope** object that represents an email header for a document.


## Syntax

_expression_.**MailEnvelope**

 _expression_ An expression that returns a **[Document](Word.Document.md)** object.


## Example

This example sets the comments for the email header of the active document.


```vb
Sub HeaderComments() 
 
 ActiveDocument.MailEnvelope.Introduction = _ 
 "Please review this document and let me know " & _ 
 "what you think. I need your input by Friday." & _ 
 " Thanks." 
 
End Sub
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]