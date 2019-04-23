---
title: Template.VBProject property (Word)
keywords: vbawd10.chm157941859
f1_keywords:
- vbawd10.chm157941859
ms.prod: word
api_name:
- Word.Template.VBProject
ms.assetid: deea632d-239f-700d-7a89-fdc0fae916ef
ms.date: 06/08/2017
localization_priority: Normal
---


# Template.VBProject property (Word)

Returns the  **VBProject** object for the specified template.


## Syntax

_expression_. `VBProject`

_expression_ Required. A variable that represents a '[Template](Word.Template.md)' object.


## Remarks

Use this property to gain access to code modules and user forms.

To view the  **VBProject** object in the object browser, you must select the **Microsoft Visual Basic for Applications Extensibility** check box in the **References** dialog box (**Tools** menu) in the Visual Basic Editor.


## Example

This example displays the name of the Visual Basic project for the Normal template.


```vb
Set normProj = NormalTemplate.VBProject 
MsgBox normProj.Name
```


## See also


[Template Object](Word.Template.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]