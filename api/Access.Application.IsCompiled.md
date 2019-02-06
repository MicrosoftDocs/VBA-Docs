---
title: Application.IsCompiled property (Access)
keywords: vbaac10.chm12567
f1_keywords:
- vbaac10.chm12567
ms.prod: access
api_name:
- Access.Application.IsCompiled
ms.assetid: c3b80c32-2aba-432c-1909-4c8172a3bebf
ms.date: 02/05/2019
localization_priority: Normal
---


# Application.IsCompiled property (Access)

The **IsCompiled** property returns a **Boolean** value indicating whether the Visual Basic project is in a compiled state. Read-only **Boolean**.


## Syntax

_expression_.**IsCompiled**

_expression_ A variable that represents an **[Application](Access.Application.md)** object.


## Remarks

The **IsCompiled** property of the **[Application](Access.Application.md)** object is **False** when the project has never been fully compiled, if a module has been added, edited, or deleted after compilation, or if a module hasn't been saved in a compiled state.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]