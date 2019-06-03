---
title: Table.Borders property (Word)
keywords: vbawd10.chm156304460
f1_keywords:
- vbawd10.chm156304460
ms.prod: word
api_name:
- Word.Table.Borders
ms.assetid: 904bce6b-db91-32be-f65d-7200f9a63be8
ms.date: 06/08/2017
localization_priority: Normal
---


# Table.Borders property (Word)

Returns a  **[Borders](Word.borders.md)** collection that represents all the borders for the specified object.


## Syntax

_expression_.**Borders**

_expression_ Required. A variable that represents a '[Table](Word.Table.md)' object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example applies inside and outside borders to the first table in the active document.


```vb
Set myTable = ActiveDocument.Tables(1) 
With myTable.Borders 
 .InsideLineStyle = wdLineStyleSingle 
 .OutsideLineStyle = wdLineStyleDouble 
End With
```


## See also


[Table Object](Word.Table.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
