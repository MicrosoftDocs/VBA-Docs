---
title: TextColumn.Width property (Word)
ms.prod: word
api_name:
- Word.TextColumn.Width
ms.assetid: 4050636e-0721-56b2-7a63-3f56906e3ca6
ms.date: 06/08/2017
localization_priority: Normal
---


# TextColumn.Width property (Word)

Returns or sets the width, in [points](../language/glossary/vbe-glossary.md#point), of the specified text columns. Read/write  **Long**.


## Syntax

_expression_.**Width**

_expression_ A variable that represents a '[TextColumn](Word.TextColumn.md)' object.


## Example

This example formats the section that includes the selection as three columns. The  **For Each** loop is used to display the width of each column in the **TextColumns** collection.


```vb
Selection.PageSetup.TextColumns.SetCount NumColumns:=3 
For Each acol In Selection.PageSetup.TextColumns 
 MsgBox "Width= " & PointsToInches(acol.Width) 
Next acol
```


## See also


[TextColumn Object](Word.TextColumn.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]