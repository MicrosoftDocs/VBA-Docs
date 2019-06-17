---
title: Selection.Document property (Word)
keywords: vbawd10.chm158663659
f1_keywords:
- vbawd10.chm158663659
ms.prod: word
api_name:
- Word.Selection.Document
ms.assetid: 03b4bfd7-8d4a-f069-0c28-41be2ead8614
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.Document property (Word)

Returns a  **[Document](Word.Document.md)** object associated with the specified selection. Read-only.


## Syntax

_expression_.**Document**

_expression_ A variable that represents a **[Selection](Word.Selection.md)** object.


## Example

This example displays the document name and path for the selection.


```vb
Msgbox Selection.Document.FullName
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]