---
title: Row.Select method (Word)
keywords: vbawd10.chm156303359
f1_keywords:
- vbawd10.chm156303359
ms.prod: word
api_name:
- Word.Row.Select
ms.assetid: f3c31e32-b316-abf2-fec6-b76e8950b1b5
ms.date: 06/08/2017
localization_priority: Normal
---


# Row.Select method (Word)

Selects the specified table row.


## Syntax

_expression_.**Select**

_expression_ Required. A variable that represents a '[Row](Word.Row.md)' object.


## Remarks

After using this method, use the **Selection** object to work with the selected row. For more information, see [Working with the Selection Object](../word/Concepts/Working-with-Word/working-with-the-selection-object.md).


## Example

This example selects row one in table one of Report.doc.


```vb
Documents("Report.doc").Tables(1).Rows(1).Select
```


## See also


[Row Object](Word.Row.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]