---
title: Document.ManualHyphenation method (Word)
keywords: vbawd10.chm158007401
f1_keywords:
- vbawd10.chm158007401
ms.prod: word
api_name:
- Word.Document.ManualHyphenation
ms.assetid: ffd4aace-f9e3-a7ef-9dab-5694891a68ab
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.ManualHyphenation method (Word)

Initiates manual hyphenation of a document, one line at a time.


## Syntax

_expression_. `ManualHyphenation`

_expression_ Required. A variable that represents a **[Document](Word.Document.md)** object.


## Remarks

When you use the **ManualHyphenation** method, Word prompts he user to accept or decline suggested hyphenations.


## Example

This example starts manual hyphenation of the active document.


```vb
ActiveDocument.ManualHyphenation
```

This example sets hyphenation options and then starts manual hyphenation of MyDoc.doc.




```vb
With Documents("MyDoc.doc") 
 .HyphenationZone = InchesToPoints(0.25) 
 .HyphenateCaps = False 
 .ManualHyphenation 
End With
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]