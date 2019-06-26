---
title: Bookmark.Delete method (Word)
keywords: vbawd10.chm157810699
f1_keywords:
- vbawd10.chm157810699
ms.prod: word
api_name:
- Word.Bookmark.Delete
ms.assetid: d5b43d2b-b605-1631-b111-9ba851d0ef1c
ms.date: 06/08/2017
localization_priority: Normal
---


# Bookmark.Delete method (Word)

Deletes the specified bookmark.


## Syntax

_expression_.**Delete**

_expression_ Required. A variable that represents a **[Bookmark](Word.Bookmark.md)** object.


## Example

If a bookmark named "temp" exists in the active document, this example deletes the bookmark.


```vb
Sub DeleteBookmark() 
 Dim intResponse As Integer 
 Dim strBookmark As String 
 
 strBookmark = "temp" 
 
 intResponse = MsgBox("Are you sure you want to delete " _ 
 & "the bookmark named """ & strBookmark & """?", vbYesNo) 
 
 If intResponse = vbYes Then 
 If ActiveDocument.Bookmarks.Exists(Name:=strBookmark) Then 
 ActiveDocument.Bookmarks(Index:=strBookmark).Delete 
 End If 
 End If 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]