---
title: CodeProject.Application property (Access)
keywords: vbaac10.chm12722
f1_keywords:
- vbaac10.chm12722
ms.prod: access
api_name:
- Access.CodeProject.Application
ms.assetid: d3ef226b-cb93-9e55-5456-c692c7615860
ms.date: 02/27/2019
localization_priority: Normal
---


# CodeProject.Application property (Access)

You can use the **Application** property to access the active Microsoft Access **[Application](Access.Application.md)** object and its related properties. Read-only **Application** object.


## Syntax

_expression_.**Application**

_expression_ A variable that represents a **[CodeProject](Access.CodeProject.md)** object.


## Remarks

The **Application** property is set by Microsoft Access and is read-only in all views.

Each Microsoft Access object has an **Application** property that returns the current **Application** object. You can use this property to access any of the object's properties. For example, you could refer to the menu bar for the **Application** object from the current form by using the following syntax.

```vb
Me.Application.MenuBar 

```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]