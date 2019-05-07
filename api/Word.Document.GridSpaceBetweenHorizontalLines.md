---
title: Document.GridSpaceBetweenHorizontalLines property (Word)
keywords: vbawd10.chm158007602
f1_keywords:
- vbawd10.chm158007602
ms.prod: word
api_name:
- Word.Document.GridSpaceBetweenHorizontalLines
ms.assetid: 79cac143-588d-d719-c653-f24852f288b6
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.GridSpaceBetweenHorizontalLines property (Word)

Returns or sets the interval at which Microsoft Word displays horizontal character gridlines in print layout view. Read/write  **Long**.


## Syntax

_expression_. `GridSpaceBetweenHorizontalLines`

_expression_ A variable that represents a **[Document](Word.Document.md)** object.


## Example

This example sets Microsoft Word to display every fifth horizontal character gridline.


```vb
ActiveDocument.GridSpaceBetweenHorizontalLines = 5
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]