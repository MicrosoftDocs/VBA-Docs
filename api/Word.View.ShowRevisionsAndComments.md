---
title: View.ShowRevisionsAndComments property (Word)
keywords: vbawd10.chm161808418
f1_keywords:
- vbawd10.chm161808418
ms.prod: word
api_name:
- Word.View.ShowRevisionsAndComments
ms.assetid: b59de20c-cff0-0621-cb0d-aa04d77f1347
ms.date: 06/08/2017
localization_priority: Normal
---


# View.ShowRevisionsAndComments property (Word)

 **True** for Microsoft Word to display revisions and comments that were made to a document with Track Changes enabled. Read/write **Boolean**.


## Syntax

_expression_. `ShowRevisionsAndComments`

 _expression_ An expression that returns a '[View](Word.View.md)' object.


## Example

This example hides the revisions and comments in a document. This example assumes that the document in the active window contains revisions made by one or more reviewers.


```vb
Sub ShowRevsComments() 
 ActiveWindow.View.ShowRevisionsAndComments = False 
End Sub
```


## See also


[View Object](Word.View.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]