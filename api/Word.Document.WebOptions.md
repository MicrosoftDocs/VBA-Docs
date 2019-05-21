---
title: Document.WebOptions property (Word)
keywords: vbawd10.chm158007626
f1_keywords:
- vbawd10.chm158007626
ms.prod: word
api_name:
- Word.Document.WebOptions
ms.assetid: 038eef42-8c57-8910-d8c1-7b9937f180c5
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.WebOptions property (Word)

Returns the  **[WebOptions](Word.WebOptions.md)** object, which contains document-level attributes used by Microsoft Word when you save a document as a webpage or open a webpage. Read-only.


## Syntax

_expression_.**WebOptions**

_expression_ A variable that represents a **[Document](Word.Document.md)** object.


## Example

This example specifies that cascading style sheets and Western document encoding be used when items in the active document are saved to a webpage.


```vb
Set objWO = ActiveDocument.WebOptions 
objWO.RelyOnCSS = True 
objWO.Encoding = msoEncodingWestern
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]