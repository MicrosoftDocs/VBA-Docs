---
title: Module.Application property (Access)
keywords: vbaac10.chm12269
f1_keywords:
- vbaac10.chm12269
ms.prod: access
api_name:
- Access.Module.Application
ms.assetid: 9237a6d4-8c68-9d58-f696-6525f42963d0
ms.date: 03/22/2019
localization_priority: Normal
---


# Module.Application property (Access)

You can use the **Application** property to access the active Microsoft Access **[Application](Access.Application.md)** object and its related properties. Read-only **Application** object.


## Syntax

_expression_.**Application**

_expression_ A variable that represents a **[Module](Access.Module.md)** object.


## Remarks

The **Application** property is set by Microsoft Access and is read-only in all views.

Each Microsoft Access object has an **Application** property that returns the current **Application** object. You can use this property to access any of the object's properties. For example, you could refer to the menu bar for the **Application** object from the current form by using the following syntax.

```vb
Me.Application.MenuBar 

```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]