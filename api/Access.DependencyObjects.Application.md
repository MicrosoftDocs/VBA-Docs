---
title: DependencyObjects.Application property (Access)
keywords: vbaac10.chm13266
f1_keywords:
- vbaac10.chm13266
ms.prod: access
api_name:
- Access.DependencyObjects.Application
ms.assetid: 39fbdeb1-bb2a-569c-7d6c-4dddf47aec51
ms.date: 03/06/2019
localization_priority: Normal
---


# DependencyObjects.Application property (Access)

You can use the **Application** property to access the active Microsoft Access **[Application](Access.Application.md)** object and its related properties. Read-only **Application** object.


## Syntax

_expression_.**Application**

_expression_ A variable that represents a **[DependencyObjects](Access.DependencyObjects.md)** object.


## Remarks

The **Application** property is set by Microsoft Access and is read-only in all views.

Each Microsoft Access object has an **Application** property that returns the current **Application** object. You can use this property to access any of the object's properties. For example, you could refer to the menu bar for the **Application** object from the current form by using the following syntax.

```vb
Me.Application.MenuBar 

```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]