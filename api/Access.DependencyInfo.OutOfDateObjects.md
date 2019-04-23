---
title: DependencyInfo.OutOfDateObjects property (Access)
keywords: vbaac10.chm13276
f1_keywords:
- vbaac10.chm13276
ms.prod: access
api_name:
- Access.DependencyInfo.OutOfDateObjects
ms.assetid: 3e6465c0-c1e4-0b26-de2e-0610e3a40273
ms.date: 03/06/2019
localization_priority: Normal
---


# DependencyInfo.OutOfDateObjects property (Access)

Returns a **[DependencyObjects](Access.DependencyObjects.md)** collection that represents the **[AccessObject](Access.AccessObject.md)** objects for which the dependency information is outdated. Read-only.


## Syntax

_expression_.**OutOfDateObjects**

_expression_ A variable that represents a **[DependencyInfo](Access.DependencyInfo.md)** object.


## Remarks

You can use the following code to update the dependency information for all of the objects in the database.

```vb
Application.CurrentProject.UpdateDependencyInfo
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]