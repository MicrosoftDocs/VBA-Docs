---
title: Image.Application property (Access)
keywords: vbaac10.chm10354
f1_keywords:
- vbaac10.chm10354
ms.prod: access
api_name:
- Access.Image.Application
ms.assetid: 7c308c10-ee19-f162-a9e4-2d6d6b9eafb0
ms.date: 03/21/2019
localization_priority: Normal
---


# Image.Application property (Access)

You can use the **Application** property to access the active Microsoft Access **[Application](Access.Application.md)** object and its related properties. Read-only **Application** object.


## Syntax

_expression_.**Application**

_expression_ A variable that represents an **[Image](Access.Image.md)** object.


## Remarks

The **Application** property is set by Microsoft Access and is read-only in all views.

Each Microsoft Access object has an **Application** property that returns the current **Application** object. You can use this property to access any of the object's properties. For example, you could refer to the menu bar for the **Application** object from the current form by using the following syntax.

```vb
Me.Application.MenuBar 

```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]