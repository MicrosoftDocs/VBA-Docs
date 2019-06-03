---
title: Document.ListTemplates property (Word)
keywords: vbawd10.chm158007359
f1_keywords:
- vbawd10.chm158007359
ms.prod: word
api_name:
- Word.Document.ListTemplates
ms.assetid: dc27553a-7083-4f14-ffd6-0f440982a79c
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.ListTemplates property (Word)

Returns a  **ListTemplates** collection that represents all the list formats for the specified document. Read-only.


## Syntax

_expression_. `ListTemplates`

_expression_ A variable that represents a **[Document](Word.Document.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md). The ListTemplates property is a member of the [Document](Word.Document.md), [ListGallery](Word.ListGallery.md), and [Template](Word.Template.md) objects.


## Example

This example displays the number of list templates used in the active document.


```vb
Msgbox ActiveDocument.ListTemplates.Count
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]