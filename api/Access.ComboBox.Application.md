---
title: ComboBox.Application property (Access)
keywords: vbaac10.chm11356
f1_keywords:
- vbaac10.chm11356
ms.prod: access
api_name:
- Access.ComboBox.Application
ms.assetid: 21c195f2-7a1f-a945-504e-6c1a7fa7f01f
ms.date: 02/28/2019
localization_priority: Normal
---


# ComboBox.Application property (Access)

You can use the **Application** property to access the active Microsoft Access **[Application](Access.Application.md)** object and its related properties. Read-only **Application** object.


## Syntax

_expression_.**Application**

_expression_ A variable that represents a **[ComboBox](Access.ComboBox.md)** object.


## Remarks

The **Application** property is set by Microsoft Access and is read-only in all views.

Each Access object has an **Application** property that returns the current **Application** object. You can use this property to access any of the object's properties. For example, you could refer to the menu bar for the **Application** object from the current form by using the following syntax.

```vb
Me.Application.MenuBar 

```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]