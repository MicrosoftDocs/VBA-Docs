---
title: Bookmark.StoryType property (Word)
keywords: vbawd10.chm157810694
f1_keywords:
- vbawd10.chm157810694
ms.prod: word
api_name:
- Word.Bookmark.StoryType
ms.assetid: 378a37f5-9ffd-1d11-4a59-b7f54f65e96b
ms.date: 06/08/2017
localization_priority: Normal
---


# Bookmark.StoryType property (Word)

Returns the story type for the specified range, selection, or bookmark. Read-only  **WdStoryType**.


## Syntax

 _expression_. `StoryType`

 _expression_ Required. A variable that represents a '[Bookmark](Word.Bookmark.md)' object.


## Example

This example selects the bookmark named "temp" if the bookmark is contained in the main story of the active document.


```vb
If ActiveDocument.Bookmarks.Exists("temp") = True Then 
 Set myBookmark = ActiveDocument.Bookmarks("temp") 
 If myBookmark.StoryType = wdMainTextStory _ 
 Then myBookmark.Select 
End If
```


## See also


[Bookmark Object](Word.Bookmark.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]