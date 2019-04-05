---
title: Application.Repeat method (Excel)
keywords: vbaxl10.chm133200
f1_keywords:
- vbaxl10.chm133200
ms.prod: excel
api_name:
- Excel.Application.Repeat
ms.assetid: ce8f6340-174e-b6cf-0f99-f39be2cde5c2
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.Repeat method (Excel)

Repeats the last user-interface action.


## Syntax

_expression_.**Repeat**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Remarks

This method repeats only the last action taken by the user before running the macro, and it must be the first line in the macro. It cannot be used to repeat Visual Basic commands.


## Example

This example repeats the last user-interface command. The example must be the first line in a macro.

```vb
Application.Repeat
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]