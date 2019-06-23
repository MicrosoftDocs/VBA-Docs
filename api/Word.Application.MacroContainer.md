---
title: Application.MacroContainer property (Word)
keywords: vbawd10.chm158335031
f1_keywords:
- vbawd10.chm158335031
ms.prod: word
api_name:
- Word.Application.MacroContainer
ms.assetid: 9c2d37b8-d5c3-d13b-3bf9-54e1352b1855
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.MacroContainer property (Word)

Returns a  **[Template](Word.Template.md)** or **[Document](Word.Document.md)** object that represents the template or document in which the module that contains the running procedure is stored.


## Syntax

_expression_. `MacroContainer`

_expression_ A variable that represents an **[Application](Word.Application.md)** object. 


## Example

This example displays the name of the document or template in which the running procedure is stored.


```vb
Set cntnr = MacroContainer 
MsgBox cntnr.Name
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]