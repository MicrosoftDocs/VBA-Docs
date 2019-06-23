---
title: Application.AutomaticChange method (Word)
keywords: vbawd10.chm158335306
f1_keywords:
- vbawd10.chm158335306
ms.prod: word
api_name:
- Word.Application.AutomaticChange
ms.assetid: 40538590-c71c-aafb-4e3b-e8759cb0116c
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.AutomaticChange method (Word)

Performs an  **AutoFormat** action when there is a change suggested by the Office Assistant. If no AutoFormat action is active, this method generates an error.


## Syntax

_expression_. `AutomaticChange`

_expression_ Required. A variable that represents an **[Application](Word.Application.md)** object. 


## Example

This example completes an Office Assistant AutoFormat action if one is active.


```vb
Application.AutomaticChange
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]