---
title: Application.GlobalTaskFilters property (Project)
keywords: vbapj.chm132265
f1_keywords:
- vbapj.chm132265
ms.prod: project-server
api_name:
- Project.Application.GlobalTaskFilters
ms.assetid: 1f85f0c7-9cb8-e531-c690-6ea795ebaa94
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.GlobalTaskFilters property (Project)

Gets or sets a  **[Filters](Project.Filter.md)** collection representing the task filters in the Global.mpt file. Read/write **Filters**.


## Syntax

_expression_. `GlobalTaskFilters`

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Remarks

 In Project Professional, you can also add a task filter to the enterprise global template. First open the enterprise global template, making it the active project, and then run the **Add** method of the **Filters** object.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]