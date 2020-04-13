---
title: Bookmark.Empty property (Word)
keywords: vbawd10.chm157810690
f1_keywords:
- vbawd10.chm157810690
ms.prod: word
api_name:
- Word.Bookmark.Empty
ms.assetid: 88675e63-9e34-e9e4-247a-3d3281bbf2e2
ms.date: 06/08/2017
localization_priority: Normal
---


# Bookmark.Empty property (Word)

 **True** if the specified bookmark is empty. Read-only **Boolean**.


## Syntax

_expression_. `Empty`

_expression_ A variable that represents a '[Bookmarks](Word.bookmarks.md)' object.


## Remarks

An empty bookmark marks a location (a collapsed selection); it doesn't mark any text. An error occurs if the specified bookmark doesn't exist. Use the **[Exists](Word.Bookmarks.Exists.md)** property to determine whether the bookmark exists.


## Example

This example determines whether the bookmark named "temp" exists and whether it is empty.


```vb
If ActiveDocument.Bookmarks.Exists("temp") = True Then 
 If ActiveDocument.Bookmarks("temp").Empty = True Then _ 
 MsgBox "The Temp bookmark is empty" 
End If
```


## See also


[Bookmark Object](Word.Bookmark.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]