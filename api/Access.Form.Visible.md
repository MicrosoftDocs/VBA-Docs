---
title: Form.Visible property (Access)
keywords: vbaac10.chm13415
f1_keywords:
- vbaac10.chm13415
ms.prod: access
api_name:
- Access.Form.Visible
ms.assetid: 85567666-297a-3380-2d08-864d44b637a1
ms.date: 02/27/2019
localization_priority: Normal
---


# Form.Visible property (Access)

Returns or sets whether the object is visible. Read/write **Boolean**.


## Syntax

_expression_.**Visible**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Remarks

To hide an object when printing, use the **DisplayWhen** property.

You can use the **Visible** property to hide a control on a form or report by including the property in a macro or event procedure that runs when the **Current** event occurs. For example, you can show or hide a congratulatory message next to a salesperson's monthly sales total in a sales report, depending on the sales total.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]