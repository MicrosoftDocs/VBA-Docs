---
title: TableOfContents.HidePageNumbersInWeb property (Word)
keywords: vbawd10.chm152240140
f1_keywords:
- vbawd10.chm152240140
ms.prod: word
api_name:
- Word.TableOfContents.HidePageNumbersInWeb
ms.assetid: 81d77980-099e-e048-b219-d10b64cd6a38
ms.date: 06/08/2017
localization_priority: Normal
---


# TableOfContents.HidePageNumbersInWeb property (Word)

Returns or sets whether page numbers in a table of contents or a table of figures should be hidden when publishing to the Web. Read/write  **Boolean**.


## Syntax

_expression_. `HidePageNumbersInWeb`

_expression_ A variable that represents a '[TableOfContents](Word.TableOfContents.md)' collection.


## Example

This example hides page numbers in the first table of contents if the document is to be published to the Web.


```vb
ActiveDocument.TableOfContents(1).HidePageNumbersInWeb = True
```


## See also


[TableOfContents Object](Word.TableOfContents.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]