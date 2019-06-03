---
title: Document.Bookmarks property (Word)
keywords: vbawd10.chm158007300
f1_keywords:
- vbawd10.chm158007300
ms.prod: word
api_name:
- Word.Document.Bookmarks
ms.assetid: 47aaace6-843c-0a2d-e584-7a8ef52f6953
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.Bookmarks property (Word)

Returns a  **[Bookmarks](Word.bookmarks.md)** collection that represents all the bookmarks in a document. Read-only.


## Syntax

_expression_. `Bookmarks`

_expression_ A variable that represents a **[Document](Word.Document.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example retrieves the starting and ending character positions for the first bookmark in the active document.


```vb
With ActiveDocument.Bookmarks(1) 
 BookStart = .Start 
 BookEnd = .End 
End With
```

This example uses the aMarks() array to store the name of each bookmark contained in the active document.




```vb
If ActiveDocument.Bookmarks.Count >= 1 Then 
 ReDim aMarks(ActiveDocument.Bookmarks.Count - 1) 
 i = 0 
 For Each aBookmark In ActiveDocument.Bookmarks 
 aMarks(i) = aBookmark.Name 
 i = i + 1 
 Next aBookmark 
End If
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]