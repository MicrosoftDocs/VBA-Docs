---
title: Application.Application property (Access)
keywords: vbaac10.chm12495
f1_keywords:
- vbaac10.chm12495
api_name:
- Access.Application.Application
ms.assetid: 2be2025d-263d-23d9-1b70-fce5108b4875
ms.date: 02/05/2019
ms.localizationpriority: medium
---


# Application.Application property (Access)

Use the **Application** property to access the active Microsoft Access **Application** object and its related properties. Read-only **Application** object.


## Syntax

_expression_.**Application**

_expression_ A variable that represents an **[Application](Access.Application.md)** object.


## Remarks

The **Application** property is set by Microsoft Access and is read-only in all views.

Each Microsoft Access object has an **Application** property that returns the current **Application** object. Use this property to access any of the object's properties. For example, you could refer to the menu bar for the **Application** object from the current form by using the following syntax.

```vb
Me.Application.MenuBar 

```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]