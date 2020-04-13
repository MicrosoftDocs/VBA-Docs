---
title: Bookmark.Start property (Word)
keywords: vbawd10.chm157810691
f1_keywords:
- vbawd10.chm157810691
ms.prod: word
api_name:
- Word.Bookmark.Start
ms.assetid: 42b44a7c-0d2a-daa6-a5ec-ef54d01cb5c3
ms.date: 06/08/2017
localization_priority: Normal
---


# Bookmark.Start property (Word)

Returns or sets the starting character position of a bookmark. Read/write  **Long**.


## Syntax

_expression_.**Start**

_expression_ A variable that represents a **[Bookmark](Word.Bookmark.md)** object.


## Remarks

If this property is set to a value larger than that of the **[End](Word.Bookmark.End.md)** property, the **End** property is set to the same value as that of **Start** property.

 Bookmark objects have starting and ending character positions. The starting position refers to the character position closest to the beginning of the story.

This property returns the starting character position relative to the beginning of the story. The main text story (**wdMainTextStory**) begins with character position 0 (zero). You can change the size of a bookmark by setting this property.


## Example

This example compares the ending position of the "temp" bookmark with the starting position of the "begin" bookmark.


```vb
Set Book1 = ActiveDocument.Bookmarks("begin") 
Set Book2 = ActiveDocument.Bookmarks("temp") 
If Book2.End > Book1.Start Then Book1.Select
```


## See also


[Bookmark Object](Word.Bookmark.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]