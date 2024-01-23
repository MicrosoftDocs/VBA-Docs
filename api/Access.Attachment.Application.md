---
title: Attachment.Application property (Access)
keywords: vbaac10.chm13903
f1_keywords:
- vbaac10.chm13903
api_name:
- Access.Attachment.Application
ms.assetid: db88250d-da59-300c-6f0c-3768c1bb8a7f
ms.date: 02/07/2019
ms.localizationpriority: medium
---


# Attachment.Application property (Access)

Use the **Application** property to access the active Microsoft Access **[Application](Access.Application.md)** object and its related properties. Read-only **Application** object.


## Syntax

_expression_.**Application**

_expression_ A variable that represents an **[Attachment](Access.Attachment.md)** object.


## Remarks

The **Application** property is set by Microsoft Access and is read-only in all views.

Each Access object has an **Application** property that returns the current **Application** object. Use this property to access any of the object's properties. For example, you could refer to the menu bar for the **Application** object from the current form by using the following syntax.

```vb
Me.Application.MenuBar 

```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]