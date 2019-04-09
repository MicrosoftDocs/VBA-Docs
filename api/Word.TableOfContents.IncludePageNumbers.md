---
title: TableOfContents.IncludePageNumbers property (Word)
keywords: vbawd10.chm152240136
f1_keywords:
- vbawd10.chm152240136
ms.prod: word
api_name:
- Word.TableOfContents.IncludePageNumbers
ms.assetid: 2097f009-ae18-70c3-3f3b-dbabcee06fa5
ms.date: 06/08/2017
localization_priority: Normal
---


# TableOfContents.IncludePageNumbers property (Word)

 **True** if page numbers are included in the table of contents. Read/write **Boolean**.


## Syntax

_expression_. `IncludePageNumbers`

_expression_ Required. A variable that represents a '[TableOfContents](Word.TableOfContents.md)' collection.


## Example

This example formats the first table of contents in the active document to include right-aligned page numbers.


```vb
If ActiveDocument.TablesOfContents.Count >= 1 Then 
 With ActiveDocument.TablesOfContents(1) 
 .IncludePageNumbers = True 
 .RightAlignPageNumbers = True 
 End With 
End If
```


## See also


[TableOfContents Object](Word.TableOfContents.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]