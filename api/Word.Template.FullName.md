---
title: Template.FullName property (Word)
keywords: vbawd10.chm157941767
f1_keywords:
- vbawd10.chm157941767
ms.prod: word
api_name:
- Word.Template.FullName
ms.assetid: 5a0d33f4-2034-22f6-a0ce-fa467dd97b86
ms.date: 06/08/2017
localization_priority: Normal
---


# Template.FullName property (Word)

Specifies the name of a template, including the drive or Web path. Read-only  **String**.


## Syntax

_expression_.**FullName**

_expression_ Required. A variable that represents a '[Template](Word.Template.md)' object.


## Remarks

Using this property is equivalent to using the **Path**, **PathSeparator**, and **Name** properties in sequence.


## Example

This example displays the path and file name of the template attached to the active document.


```vb
Sub TemplateName() 
 MsgBox ActiveDocument.AttachedTemplate.FullName 
End Sub
```


## See also


[Template Object](Word.Template.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]