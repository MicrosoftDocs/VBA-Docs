---
title: Field.LinkFormat property (Word)
keywords: vbawd10.chm154075146
f1_keywords:
- vbawd10.chm154075146
ms.prod: word
api_name:
- Word.Field.LinkFormat
ms.assetid: c30a1be2-0560-48e1-9103-07050157fe50
ms.date: 06/08/2017
localization_priority: Normal
---


# Field.LinkFormat property (Word)

Returns a  **LinkFormat** object that represents the link options of the specified field. Read/only.


## Syntax

_expression_. `LinkFormat`

_expression_ A variable that represents a '[Field](Word.Field.md)' object.


## Example

This example updates any fields in the active document that aren't updated automatically.


```vb
For Each afield In ActiveDocument.Fields 
 If afield.LinkFormat.AutoUpdate = False _ 
 Then afield.LinkFormat.Update 
Next afield
```


## See also


[Field Object](Word.Field.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]