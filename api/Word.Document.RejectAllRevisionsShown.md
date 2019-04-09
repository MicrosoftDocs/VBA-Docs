---
title: Document.RejectAllRevisionsShown method (Word)
keywords: vbawd10.chm158007669
f1_keywords:
- vbawd10.chm158007669
ms.prod: word
api_name:
- Word.Document.RejectAllRevisionsShown
ms.assetid: 87b46681-dbc9-e38b-e20d-5da2a9a0456f
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.RejectAllRevisionsShown method (Word)

Rejects all revisions in a document that are displayed on the screen.


## Syntax

_expression_. `RejectAllRevisionsShown`

_expression_ Required. A variable that represents a **[Document](Word.Document.md)** object.


## Example

This example hides revisions made by Jeff Smith and rejects all remaining revisions that are displayed.


```vb
Sub RejectAllChanges() 
 Dim rev As Reviewer 
 With ActiveWindow.View 
 'Show all revisions in the document 
 .ShowRevisionsAndComments = True 
 .ShowFormatChanges = True 
 .ShowInsertionsAndDeletions = True 
 
 For Each rev In .Reviewers 
 rev.Visible = True 
 Next 
 
 'Hide revisions made by "Jeff Smith" 
 .Reviewers(Index:="Jeff Smith").Visible = False 
 End With 
 
 'Reject all revisions displayed in the active view 
 ActiveDocument.RejectAllRevisionsShown 
End Sub
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]