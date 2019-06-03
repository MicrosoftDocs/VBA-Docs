---
title: Rows.Borders property (Word)
keywords: vbawd10.chm155976780
f1_keywords:
- vbawd10.chm155976780
ms.prod: word
api_name:
- Word.Rows.Borders
ms.assetid: 4c251987-5bbb-bfdb-d90f-861838f1b59d
ms.date: 06/08/2017
localization_priority: Normal
---


# Rows.Borders property (Word)

Returns a  **[Borders](Word.borders.md)** collection that represents all the borders for the specified object.


## Syntax

_expression_.**Borders**

_expression_ Required. A variable that represents a **[Rows](Word.Rows.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).




## Example

This example applies inside and outside borders to the rows in the first table in the active document.


```vb
Set myTable = ActiveDocument.Tables(1) 
With myTable.Rows.Borders 
 .InsideLineStyle = wdLineStyleSingle 
 .OutsideLineStyle = wdLineStyleDouble 
End With
```


## See also


[Rows Collection Object](Word.rows.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]