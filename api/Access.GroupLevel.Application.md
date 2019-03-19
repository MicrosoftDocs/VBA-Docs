---
title: GroupLevel.Application property (Access)
keywords: vbaac10.chm12236
f1_keywords:
- vbaac10.chm12236
ms.prod: access
api_name:
- Access.GroupLevel.Application
ms.assetid: e48bc896-a4fd-31f7-b75a-f9d44111d5e1
ms.date: 03/20/2019
localization_priority: Normal
---


# GroupLevel.Application property (Access)

You can use the **Application** property to access the active Microsoft Access **[Application](Access.Application.md)** object and its related properties. Read-only **Application** object.


## Syntax

_expression_.**Application**

_expression_ A variable that represents a **[GroupLevel](Access.GroupLevel.md)** object.


## Remarks

The **Application** property is set by Microsoft Access and is read-only in all views.

Each Microsoft Access object has an **Application** property that returns the current **Application** object. You can use this property to access any of the object's properties. For example, you could refer to the menu bar for the **Application** object from the current form by using the following syntax.

```vb
Me.Application.MenuBar 

```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]