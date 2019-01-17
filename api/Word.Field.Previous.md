---
title: Field.Previous property (Word)
keywords: vbawd10.chm154075143
f1_keywords:
- vbawd10.chm154075143
ms.prod: word
api_name:
- Word.Field.Previous
ms.assetid: be39b806-a82a-de52-190d-3f537e063d6a
ms.date: 06/08/2017
localization_priority: Normal
---


# Field.Previous property (Word)

Returns the previous object in the collection. Read-only.


## Syntax

 _expression_. `Previous`

 _expression_ A variable that represents a '[Field](Word.Field.md)' object.


## Example

This example displays the field code of the second-to-last field in the active document.


```vb
Set aField = ActiveDocument _ 
 .Fields(ActiveDocument.Fields.Count).Previous 
MsgBox "Field code = " & aField.Code
```


## See also


[Field Object](Word.Field.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]