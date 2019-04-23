---
title: BoundObjectFrame.Application property (Access)
keywords: vbaac10.chm10896
f1_keywords:
- vbaac10.chm10896
ms.prod: access
api_name:
- Access.BoundObjectFrame.Application
ms.assetid: 05b5b479-fe8b-6d03-b8de-59afa7a587b9
ms.date: 02/08/2019
localization_priority: Normal
---


# BoundObjectFrame.Application property (Access)

You can use the **Application** property to access the active Microsoft Access **[Application](Access.Application.md)** object and its related properties. Read-only **Application** object.


## Syntax

_expression_.**Application**

_expression_ A variable that represents a **[BoundObjectFrame](Access.BoundObjectFrame.md)** object.


## Remarks

The **Application** property is set by Microsoft Access and is read-only in all views.

Each Microsoft Access object has an **Application** property that returns the current **Application** object. You can use this property to access any of the object's properties. For example, you could refer to the menu bar for the **Application** object from the current form by using the following syntax.

```vb
Me.Application.MenuBar 

```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]