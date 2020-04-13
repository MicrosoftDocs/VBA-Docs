---
title: Selection.InsertColumns method (Word)
keywords: vbawd10.chm158663185
f1_keywords:
- vbawd10.chm158663185
ms.prod: word
api_name:
- Word.Selection.InsertColumns
ms.assetid: d58691b4-afa5-959a-a6a8-f202723df9f1
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.InsertColumns method (Word)

Inserts columns to the left of the column that contains the selection.


## Syntax

_expression_. `InsertColumns`

_expression_ Required. A variable that represents a **[Selection](Word.Selection.md)** object.


## Remarks

The number of columns inserted is equal to the number of columns selected. You can also insert columns by using the **[Add](Word.Columns.Add.md)** method of the **Columns** object.

If the selection isn't in a table, an error occurs.


## Example

This example inserts new columns to the left of the column that contains the selection. The number of columns inserted is equal to the number of columns selected.


```vb
If Selection.Information(wdWithInTable) = True Then 
 With Selection 
 .InsertColumns 
 .Shading.Texture = wdTexture10Percent 
 End With 
End If
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]