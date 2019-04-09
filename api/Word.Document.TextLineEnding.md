---
title: Document.TextLineEnding property (Word)
keywords: vbawd10.chm158007654
f1_keywords:
- vbawd10.chm158007654
ms.prod: word
api_name:
- Word.Document.TextLineEnding
ms.assetid: 6e1f2243-473c-0294-623e-c09588645ee3
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.TextLineEnding property (Word)

Returns or sets a  **WdLineEndingType** constant indicating how Microsoft Word marks the line and paragraph breaks in documents saved as text files. Read/write.


## Syntax

_expression_. `TextLineEnding`

_expression_ Required. A variable that represents a **[Document](Word.Document.md)** object.


## Example

This example sets the active document to enter a carriage return for line and paragraph breaks when it is saved as a text file.


```vb
Sub LineEndings() 
 ActiveDocument.TextLineEnding = wdCROnly 
End Sub
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]