---
title: TableOfContents.UpdatePageNumbers method (Word)
keywords: vbawd10.chm152240229
f1_keywords:
- vbawd10.chm152240229
ms.prod: word
api_name:
- Word.TableOfContents.UpdatePageNumbers
ms.assetid: 3b7e3080-c2bb-0a4b-2062-f1a774eeb715
ms.date: 06/08/2017
localization_priority: Normal
---


# TableOfContents.UpdatePageNumbers method (Word)

Updates the page numbers for items in the specified table of contents.


## Syntax

_expression_. `UpdatePageNumbers`

_expression_ Required. A variable that represents a '[TableOfContents](Word.TableOfContents.md)' collection.


## Example

This example inserts a page break at the insertion point and then updates the page numbers for the first table of contents in the active document.


```vb
Selection.Collapse Direction:=wdCollapseStart 
Selection.InsertBreak Type:=wdPageBreak 
ActiveDocument.TablesOfContents(1).UpdatePageNumbers
```


## See also


[TableOfContents Object](Word.TableOfContents.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]