---
title: Controls.Application property (Access)
keywords: vbaac10.chm10177
f1_keywords:
- vbaac10.chm10177
ms.prod: access
api_name:
- Access.Controls.Application
ms.assetid: c8650732-ffee-830b-9d9d-571a09af3a4c
ms.date: 06/08/2017
localization_priority: Normal
---


# Controls.Application property (Access)

You can use the  **Application** property to access the active Microsoft Access **[Application](Access.Application.md)** object and its related properties. Read-only **Application** object.


## Syntax

_expression_. `Application`

_expression_ A variable that represents a [Controls](Access.Controls.md) object.


## Remarks

The  **Application** property is set by Microsoft Access and is read-only in all views.

Each Microsoft Access object has an  **Application** property that returns the current **Application** object. You can use this property to access any of the object's properties. For example, you could refer to the menu bar for the **Application** object from the current form by using the following syntax:




```vb
Me.Application.MenuBar 

```


## See also


[Controls Collection](Access.Controls.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]