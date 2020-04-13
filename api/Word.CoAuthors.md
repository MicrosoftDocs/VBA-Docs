---
title: CoAuthors object (Word)
ms.prod: word
api_name:
- Word.CoAuthors
ms.assetid: 47fc864d-5f1b-b113-85b5-6e8b1b75c225
ms.date: 06/08/2017
localization_priority: Normal
---


# CoAuthors object (Word)

A collection of all the **[CoAuthor](Word.CoAuthor.md)** objects in the document.


## Remarks

The **CoAuthors** collection contains all the co authors in the document (authors that are actively editing the document).


## Example

The following code example gets the number of co authors in the active document.


```vb
Dim i As Integer 
 
i = ActiveDocument.CoAuthoring.Authors.Count 
 
MsgBox "The number of co authors is " & i
```


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]