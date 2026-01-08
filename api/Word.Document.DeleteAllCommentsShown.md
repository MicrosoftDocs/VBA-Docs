---
title: Document.DeleteAllCommentsShown method (Word)
keywords: vbawd10.chm158007670
f1_keywords:
- vbawd10.chm158007670
api_name:
- Word.Document.DeleteAllCommentsShown
ms.assetid: b0cdbc8e-973c-1921-a646-d2f5ef091ce9
ms.date: 06/20/2017
description: Document.DeleteAllCommentsShown method (Word)
ms.localizationpriority: medium
---


# Document.DeleteAllCommentsShown method (Word)

Deletes all revisions in a specified document that are displayed on the screen.

> [!IMPORTANT]
> `DeleteAllCommentsShown` deletes the visible subset of comments. This includes active comments (shown) as well as resolved comments (indicated by comment hint).

## Syntax

_expression_.**DeleteAllCommentsShown**

_expression_ Required. A variable that represents a **[Document](Word.Document.md)** object.


## Example

This example hides all comments made by "Jeff Smith" and deletes all other displayed comments.


```vb
Sub HideDeleteComments() 
 Dim rev As Reviewer 
 With ActiveWindow.View 
 'Display all comments and revisions 
 .ShowRevisionsAndComments = True 
 .ShowFormatChanges = True 
 .ShowInsertionsAndDeletions = True 
 
 For Each rev In .Reviewers 
 rev.Visible = True 
 Next 
 
 'Hide only the revisions/comments made by the 
 'reviewer named "Jeff Smith" 
 .Reviewers(Index:="Jeff Smith").Visible = False 
 End With 
 
 'Delete all comments displayed in the active view 
 ActiveDocument.DeleteAllCommentsShown 
End Sub
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
