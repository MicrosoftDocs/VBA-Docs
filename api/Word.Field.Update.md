---
title: Field.Update method (Word)
keywords: vbawd10.chm154075237
f1_keywords:
- vbawd10.chm154075237
ms.prod: word
api_name:
- Word.Field.Update
ms.assetid: e4e941aa-3223-ae0b-8366-9e14d92fff52
ms.date: 06/08/2017
localization_priority: Normal
---


# Field.Update method (Word)

Updates the result of the field. Returns  **True** if the field is updated successfully.


## Syntax

_expression_.**Update**

_expression_ Required. A variable that represents a '[Field](Word.Field.md)' object.


## Return value

Boolean


## Example

This example updates the first field in the active document. A return value of 1 (True) indicates that the fields were updated without error.


```vb
If ActiveDocument.Fields(0).Update = 1 Then 
 MsgBox "Update Successful" 
Else 
 MsgBox "Field " & ActiveDocument.Fields(0).Update & _ 
 " has an error" 
End If
```


## See also


[Field Object](Word.Field.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]