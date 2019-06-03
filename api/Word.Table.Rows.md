---
title: Table.Rows property (Word)
keywords: vbawd10.chm156303461
f1_keywords:
- vbawd10.chm156303461
ms.prod: word
api_name:
- Word.Table.Rows
ms.assetid: e4cc7541-15fe-97b6-0fe6-90d561a85420
ms.date: 06/08/2017
localization_priority: Normal
---


# Table.Rows property (Word)

Returns a  **Rows** collection that represents all the table rows within a table. Read-only.


## Syntax

_expression_.**Rows**

_expression_ A variable that represents a '[Table](Word.Table.md)' object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example deletes the second row from the first table in the active document.

```vb
ActiveDocument.Tables(1).Rows(2).Delete
```


## See also


[Table Object](Word.Table.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]