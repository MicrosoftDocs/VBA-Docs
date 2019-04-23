---
title: FormField.Result property (Word)
keywords: vbawd10.chm153616394
f1_keywords:
- vbawd10.chm153616394
ms.prod: word
api_name:
- Word.FormField.Result
ms.assetid: b1e242d0-11d1-4b85-28b2-6fc821ed3c96
ms.date: 06/08/2017
localization_priority: Normal
---


# FormField.Result property (Word)

Returns a  **String** that represents the result of the specified form field. Read/write.


## Syntax

_expression_. `Result`

_expression_ Required. A variable that represents a '[FormField](Word.FormField.md)' object.


## Example

This example displays the result of each form field in the active document.


```vb
For Each aField In ActiveDocument.FormFields 
 MsgBox aField.Result 
Next aField
```


## See also


[FormField Object](Word.FormField.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]