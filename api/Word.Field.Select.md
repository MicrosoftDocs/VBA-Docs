---
title: Field.Select method (Word)
keywords: vbawd10.chm154140671
f1_keywords:
- vbawd10.chm154140671
ms.prod: word
api_name:
- Word.Field.Select
ms.assetid: 03fa304c-acc7-30a5-7dfa-06098bbdac7a
ms.date: 06/08/2017
localization_priority: Normal
---


# Field.Select method (Word)

Selects the specified field.


## Syntax

_expression_.**Select**

_expression_ Required. A variable that represents a '[Field](Word.Field.md)' object.


## Remarks

After using this method, use the **[Selection](Word.Selection.md)** object to work with the selected items. For more information, see [Working with the Selection Object](../word/Concepts/Working-with-Word/working-with-the-selection-object.md).


## Example

This example updates and selects the first field in the active document.


```vb
ActiveDocument.ActiveWindow.View.FieldShading = _ 
 wdFieldShadingWhenSelected 
If ActiveDocument.Fields.Count >= 1 Then 
 With ActiveDocument.Fields(1) 
 .Update 
 .Select 
 End With 
End If
```


## See also


[Field Object](Word.Field.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]