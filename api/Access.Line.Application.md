---
title: Line.Application property (Access)
keywords: vbaac10.chm10322
f1_keywords:
- vbaac10.chm10322
ms.prod: access
api_name:
- Access.Line.Application
ms.assetid: d12619b5-99ad-f3ff-9d28-19cd9991d749
ms.date: 03/22/2019
localization_priority: Normal
---


# Line.Application property (Access)

You can use the **Application** property to access the active Microsoft Access **[Application](Access.Application.md)** object and its related properties. Read-only **Application** object.


## Syntax

_expression_.**Application**

_expression_ A variable that represents a **[Line](Access.Line.md)** object.


## Remarks

The **Application** property is set by Microsoft Access and is read-only in all views.

Each Microsoft Access object has an **Application** property that returns the current **Application** object. You can use this property to access any of the object's properties. For example, you could refer to the menu bar for the **Application** object from the current form by using the following syntax.

```vb
Me.Application.MenuBar 

```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]