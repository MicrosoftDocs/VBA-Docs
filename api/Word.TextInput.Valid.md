---
title: TextInput.Valid property (Word)
keywords: vbawd10.chm153550848
f1_keywords:
- vbawd10.chm153550848
ms.prod: word
api_name:
- Word.TextInput.Valid
ms.assetid: cf8399fd-d69e-6a49-dcbc-1b548ebc9002
ms.date: 06/08/2017
localization_priority: Normal
---


# TextInput.Valid property (Word)

 **True** if the specified form field object is a valid check box form field. Read-only **Boolean**. .


## Syntax

_expression_. `Valid`

_expression_ A variable that represents a '[TextInput](Word.TextInput.md)' object.


## Example

This example determines whether the first form field in the active document is a text form field. If the **Valid** property is **True**, the contents of the text form field are changed to "Hello."


```vb
If ActiveDocument.FormFields(1).TextInput.Valid = True Then 
 ActiveDocument.FormFields(1).Result = "Hello" 
End If
```


## See also


[TextInput Object](Word.TextInput.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]