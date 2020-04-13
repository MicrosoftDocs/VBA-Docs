---
title: Bookmarks.DefaultSorting property (Word)
keywords: vbawd10.chm157745155
f1_keywords:
- vbawd10.chm157745155
ms.prod: word
api_name:
- Word.Bookmarks.DefaultSorting
ms.assetid: 86f17298-8a11-a5d6-05fd-4cb87f6e5f91
ms.date: 06/08/2017
localization_priority: Normal
---


# Bookmarks.DefaultSorting property (Word)

Returns or sets the sorting option for bookmark names displayed in the **Bookmark** dialog box (**Insert** menu). Read/write **WdBookmarkSortBy**.


## Syntax

_expression_. `DefaultSorting`

_expression_ Required. A variable that represents a '[Bookmarks](Word.bookmarks.md)' collection.


## Remarks

This property doesn't affect the order of  **Bookmark** objects in the **Bookmarks** collection.


## Example

This example sorts bookmarks by location and then displays the Bookmark dialog box.


```vb
ActiveDocument.Bookmarks.DefaultSorting = wdSortByLocation 
Dialogs(wdDialogInsertBookmark).Show
```


## See also


[Bookmarks Collection Object](Word.bookmarks.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]