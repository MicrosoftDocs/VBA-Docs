---
title: WebServices.Application property (Access)
keywords: vbaac10.chm14551
f1_keywords:
- vbaac10.chm14551
ms.prod: access
api_name:
- Access.WebServices.Application
ms.assetid: c6973a91-7dcf-633e-ce3f-d494d5614e7d
ms.date: 03/26/2019
localization_priority: Normal
---


# WebServices.Application property (Access)

You can use the **Application** property to access the active Microsoft Access **[Application](Access.Application.md)** object and its related properties. Read-only **Application** object.


## Syntax

_expression_.**Application**

_expression_ A variable that represents a **[WebServices](Access.WebServices.md)** object.


## Remarks

The **Application** property is set by Microsoft Access and is read-only in all views.

Each Microsoft Access object has an **Application** property that returns the current **Application** object. You can use this property to access any of the object's properties. For example, you could refer to the menu bar for the **Application** object from the current form by using the following syntax.

```vb
Me.Application.MenuBar 

```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]