---
title: EmptyCell.Application property (Access)
keywords: vbaac10.chm14297
f1_keywords:
- vbaac10.chm14297
api_name:
- Access.EmptyCell.Application
ms.assetid: df8b9d6b-3065-ac43-3ead-ff504bd76db1
ms.date: 03/08/2019
ms.localizationpriority: medium
---


# EmptyCell.Application property (Access)

Use the **Application** property to access the active Microsoft Access **[Application](Access.Application.md)** object and its related properties. Read-only **Application** object.


## Syntax

_expression_.**Application**

_expression_ A variable that represents an **[EmptyCell](Access.EmptyCell.md)** object.


## Remarks

The **Application** property is set by Microsoft Access and is read-only in all views.

Each Microsoft Access object has an **Application** property that returns the current **Application** object. Use this property to access any of the object's properties. For example, you could refer to the menu bar for the **Application** object from the current form by using the following syntax.

```vb
Me.Application.MenuBar 

```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]