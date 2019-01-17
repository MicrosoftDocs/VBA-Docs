---
title: Document.StyleSheets property (Word)
keywords: vbawd10.chm158007656
f1_keywords:
- vbawd10.chm158007656
ms.prod: word
api_name:
- Word.Document.StyleSheets
ms.assetid: 119a2ecb-9cbd-531e-2145-fc28da798a05
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.StyleSheets property (Word)

Returns a  **[StyleSheets](Word.StyleSheets.md)** collection that represents the Web style sheets attached to a document.


## Syntax

 _expression_. `StyleSheets`

 _expression_ An variable that represents a '[Document](Word.Document.md)' object.


## Example

This example adds a style sheet to the active document and places it highest in the list of style sheets attached to the document. This example assumes that you have a style sheet document named "Website.css" located on your drive C.


```vb
Sub Styshts() 
 ActiveDocument.StyleSheets.Add _ 
 FileName:="c:\Website.css", _ 
 Precedence:=wdStyleSheetPrecedenceHighest 
End Sub
```


## See also


[Document Object](Word.Document.md)

