---
title: Document.ReadOnly property (Word)
keywords: vbawd10.chm158007340
f1_keywords:
- vbawd10.chm158007340
ms.prod: word
api_name:
- Word.Document.ReadOnly
ms.assetid: 57421a93-808f-f216-5110-0c3b80cf6e04
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.ReadOnly property (Word)

 **True** if changes to the document cannot be saved to the original document. Read-only **Boolean**.


## Syntax

 _expression_. `ReadOnly`

 _expression_ Required. A variable that represents a '[Document](Word.Document.md)' object.


## Example

This example saves the active document if it isn't read-only.


```vb
If ActiveDocument.ReadOnly = False Then ActiveDocument.Save
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]