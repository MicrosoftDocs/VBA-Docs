---
title: Application.Visible property (Access)
keywords: vbaac10.chm12513
f1_keywords:
- vbaac10.chm12513
ms.prod: access
api_name:
- Access.Application.Visible
ms.assetid: ac1558c1-68c4-fdf1-4f59-77343b7b5e59
ms.date: 02/05/2019
localization_priority: Normal
---


# Application.Visible property (Access)

Returns or sets whether a Microsoft Access application is minimized. Read/write **Boolean**.


## Syntax

_expression_.**Visible**

_expression_ A variable that represents an **[Application](Access.Application.md)** object.


## Remarks

When an application is launched by the user, the **Visible** and **[UserControl](Access.Application.UserControl.md)** properties of the **Application** object are both set to **True**. When the **UserControl** property is set to **True**, it isn't possible to set the **Visible** property of the object to **False**.

When an **Application** object is created by using Automation, the **Visible** and **UserControl** properties of the object are both set to **False**.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]