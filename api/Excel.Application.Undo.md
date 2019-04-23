---
title: Application.Undo method (Excel)
keywords: vbaxl10.chm133221
f1_keywords:
- vbaxl10.chm133221
ms.prod: excel
api_name:
- Excel.Application.Undo
ms.assetid: b56bb8a0-2cd1-356a-03ba-47eb6f56f455
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.Undo method (Excel)

Cancels the last user-interface action.


## Syntax

_expression_.**Undo**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Remarks

This method undoes only the last action taken by the user before running the macro, and it must be the first line in the macro. It cannot be used to undo Visual Basic commands.


## Example

This example cancels the last user-interface action. The example must be the first line in a macro.


```vb
Application.Undo
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
