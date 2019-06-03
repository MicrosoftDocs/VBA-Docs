---
title: Selection.Rows property (Word)
keywords: vbawd10.chm158662959
f1_keywords:
- vbawd10.chm158662959
ms.prod: word
api_name:
- Word.Selection.Rows
ms.assetid: 800edca7-fc0f-ed73-ae3a-400eadcccf8b
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.Rows property (Word)

Returns a  **[Rows](Word.rows.md)** collection that represents all the table rows in a range, selection, or table. Read-only.


## Syntax

_expression_.**Rows**

_expression_ A variable that represents a **[Selection](Word.Selection.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example places a border around the cells in the row that contains the insertion point.


```vb
Selection.Collapse Direction:=wdCollapseStart 
If Selection.Information(wdWithInTable) = True Then 
 Selection.Rows(1).Borders.OutsideLineStyle = wdLineStyleSingle 
Else 
 MsgBox "The insertion point is not in a table." 
End If
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]