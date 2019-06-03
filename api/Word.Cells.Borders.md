---
title: Cells.Borders property (Word)
keywords: vbawd10.chm155845708
f1_keywords:
- vbawd10.chm155845708
ms.prod: word
api_name:
- Word.Cells.Borders
ms.assetid: df873357-9474-8f69-ae71-6df5859cbf93
ms.date: 06/08/2017
localization_priority: Normal
---


# Cells.Borders property (Word)

Returns a  **[Borders](Word.borders.md)** collection that represents all the borders for the specified object.


## Syntax

_expression_.**Borders**

_expression_ A variable that represents a '[Cells](Word.cells.md)' object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example applies inside and outside borders to the cells in the first row of the first table in the active document.


```vb
Dim objTable As Table 
Set objTable = ActiveDocument.Tables(1) 
With objTable.Rows(1).Cells.Borders 
 .InsideLineStyle = wdLineStyleSingle 
 .OutsideLineStyle = wdLineStyleDouble 
End With
```


## See also


[Cells Collection Object](Word.cells.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]