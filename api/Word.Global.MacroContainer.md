---
title: Global.MacroContainer property (Word)
keywords: vbawd10.chm163119159
f1_keywords:
- vbawd10.chm163119159
ms.prod: word
api_name:
- Word.Global.MacroContainer
ms.assetid: 9718527c-eebd-4d62-f753-da449034b404
ms.date: 06/08/2017
localization_priority: Normal
---


# Global.MacroContainer property (Word)

Returns a  **Template** or **Document** object that represents the template or document in which the module that contains the running procedure is stored.


## Syntax

_expression_. `MacroContainer`

_expression_ Required. A variable that represents a '[Global](Word.Global.md)' object.


## Example

This example displays the name of the document or template in which the running procedure is stored.


```vb
Set cntnr = MacroContainer 
MsgBox cntnr.Name
```


## See also


[Global Object](Word.Global.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]