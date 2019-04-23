---
title: TableOfFigures.IncludePageNumbers property (Word)
keywords: vbawd10.chm153157639
f1_keywords:
- vbawd10.chm153157639
ms.prod: word
api_name:
- Word.TableOfFigures.IncludePageNumbers
ms.assetid: cc363160-c1bd-b6a2-75e0-4c201db57ded
ms.date: 06/08/2017
localization_priority: Normal
---


# TableOfFigures.IncludePageNumbers property (Word)

 **True** if page numbers are included in the table of figures. Read/write **Boolean**.


## Syntax

_expression_. `IncludePageNumbers`

_expression_ Required. A variable that represents a '[TableOfFigures](Word.TableOfFigures.md)' collection.


## Example

This example formats the first table of figures in the active document to include right-aligned page numbers.


```vb
If ActiveDocument.TablesOfFigures.Count >= 1 Then 
 With ActiveDocument.TablesOfFigures(1) 
 .IncludePageNumbers = True 
 .RightAlignPageNumbers = True 
 End With 
End If
```


## See also


[TableOfFigures Object](Word.TableOfFigures.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]