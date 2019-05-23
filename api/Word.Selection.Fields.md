---
title: Selection.Fields property (Word)
keywords: vbawd10.chm158662720
f1_keywords:
- vbawd10.chm158662720
ms.prod: word
api_name:
- Word.Selection.Fields
ms.assetid: 15060502-c0cf-1c94-93ba-0db0bb045c66
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.Fields property (Word)

Returns a read-only  **[Fields](Word.fields.md)** collection that represents all the fields in the selection.


## Syntax

_expression_. `Fields`

_expression_ Required. A variable that represents a **[Selection](Word.Selection.md)** object.


## Example

This example adds a DATE field at the insertion point.


```vb
With Selection 
 .Collapse Direction:=wdCollapseStart 
 .Fields.Add Range:=Selection.Range, Type:=wdFieldDate 
End With
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]