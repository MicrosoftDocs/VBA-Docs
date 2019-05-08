---
title: RevisionsFilter.Markup property (Word)
ms.prod: word
ms.assetid: 90b90dd8-ead3-8e3c-f27e-a4614d12798c
ms.date: 06/08/2017
localization_priority: Normal
---


# RevisionsFilter.Markup property (Word)

Returns or sets a [WdRevisionsMarkup](Word.wdrevisionsmarkup.md) constant that specifies the extent of reviewer markup displayed in the document. Read/write.


## Syntax

_expression_. `Markup`

_expression_ A variable that represents a 'RevisionsFilter' object.


## Example

This example shows how to display all revisions and markup in the document in the active window. This example assumes that the document in the active window contains revisions made by one or more reviewers.


```vb
Public Sub Markup_Example()

    ActiveWindow.View.RevisionsFilter.Markup = wdRevisionsMarkupAll

End Sub
```


## Property value

 **WDREVISIONSMARKUP**


## See also


[RevisionsFilter Object](Word.revisionsfilter.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]