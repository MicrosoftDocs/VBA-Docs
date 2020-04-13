---
title: Exception.MonthPosition property (Project)
ms.prod: project-server
api_name:
- Project.Exception.MonthPosition
ms.assetid: afe3c243-5b4d-1e10-cd07-2f36f2447ba5
ms.date: 06/08/2017
localization_priority: Normal
---


# Exception.MonthPosition property (Project)

Gets or sets the position of the exception in the month, for a monthly or yearly calendar exception. Read/write  **PjExceptionPosition**.


## Syntax

_expression_. `MonthPosition`

_expression_ A variable that represents an [Exception](./Project.Exception.md) object.


## Remarks

The **MonthPosition** property can be one of the following **[PjExceptionPosition](Project.PjExceptionPosition.md)** constants: **pjFirst**, **pjSecond**, **pjThird**, **pjFourth**, **pjLast**. For example, if a monthly calendar exception is set for the second Wednesday every month, the value of **MonthPosition** is **pjSecond**.


## See also


[Exception Object](Project.Exception.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]