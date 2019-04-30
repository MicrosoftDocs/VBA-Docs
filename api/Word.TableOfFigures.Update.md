---
title: TableOfFigures.Update method (Word)
keywords: vbawd10.chm153157734
f1_keywords:
- vbawd10.chm153157734
ms.prod: word
api_name:
- Word.TableOfFigures.Update
ms.assetid: bab9ec6b-402d-a4d4-720f-b77fd187f95f
ms.date: 06/08/2017
localization_priority: Normal
---


# TableOfFigures.Update method (Word)

Updates the entries shown in a table of figures.


## Syntax

_expression_.**Update**

_expression_ Required. A variable that represents a '[TableOfFigures](Word.TableOfFigures.md)' collection.


## Remarks

 Use the **UpdatePageNumbers** method to update the page numbers of items in a table of figures.


## Example

This example updates the first table of figures in the active document.


```vb
If ActiveDocument.TablesOfFigures.Count >= 1 Then 
 ActiveDocument.TableOfFigures(1).Update 
End If
```


## See also


[TableOfFigures Object](Word.TableOfFigures.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]