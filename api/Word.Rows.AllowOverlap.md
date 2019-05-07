---
title: Rows.AllowOverlap property (Word)
keywords: vbawd10.chm155975702
f1_keywords:
- vbawd10.chm155975702
ms.prod: word
api_name:
- Word.Rows.AllowOverlap
ms.assetid: 2a5205d6-dd9c-6c12-38a3-37633cfd644b
ms.date: 06/08/2017
localization_priority: Normal
---


# Rows.AllowOverlap property (Word)

Returns or sets a value that specifies whether the specified rows can overlap other rows.


## Syntax

_expression_. `AllowOverlap`

_expression_ A variable that represents a **[Rows](Word.Rows.md)** object.


## Remarks

This property returns  **wdUndefined** if the specified rows include both overlapping rows and nonoverlapping rows. Can be set to either **True** or **False**. Read/write **Long**. Setting **AllowOverlap** to **True** also sets **WrapAroundText** to **True**, and setting **WrapAroundText** to **False** also sets **AllowOverlap** to **False**.

Because HTML doesn't support overlapping tables or shapes,  **AllowOverlap** is ignored in web layout view.


## Example

This example specifies that text wraps around the selected table and that the table doesn't overlap any other wrapped tables.


```vb
Selection.Rows.WrapAroundText = True 
Selection.Rows.AllowOverlap = False
```

This example specifies that the first shape in the active document can overlap other shapes.




```vb
ActiveDocument.Shapes(1).WrapFormat.AllowOverlap = True
```


## See also


[Rows Collection Object](Word.rows.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]