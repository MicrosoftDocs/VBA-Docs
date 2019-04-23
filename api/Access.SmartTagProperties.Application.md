---
title: SmartTagProperties.Application property (Access)
keywords: vbaac10.chm13310
f1_keywords:
- vbaac10.chm13310
ms.prod: access
api_name:
- Access.SmartTagProperties.Application
ms.assetid: 4a282407-1dc4-1a21-41b3-f7601eb59dfc
ms.date: 03/26/2019
localization_priority: Normal
---


# SmartTagProperties.Application property (Access)

You can use the **Application** property to access the active Microsoft Access **[Application](Access.Application.md)** object and its related properties. Read-only **Application** object.


## Syntax

_expression_.**Application**

_expression_ A variable that represents a **[SmartTagProperties](Access.SmartTagProperties.md)** object.


## Remarks

The **Application** property is set by Microsoft Access and is read-only in all views.

Each Microsoft Access object has an **Application** property that returns the current **Application** object. You can use this property to access any of the object's properties. For example, you could refer to the menu bar for the **Application** object from the current form by using the following syntax.

```vb
Me.Application.MenuBar 

```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]