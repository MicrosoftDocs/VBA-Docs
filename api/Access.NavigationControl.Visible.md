---
title: NavigationControl.Visible property (Access)
keywords: vbaac10.chm11056
f1_keywords:
- vbaac10.chm11056
ms.prod: access
api_name:
- Access.NavigationControl.Visible
ms.assetid: 91ea0e8c-63d1-3ca7-7f26-748f1651a1c6
ms.date: 06/08/2017
localization_priority: Normal
---


# NavigationControl.Visible property (Access)

Returns or sets whether the object is visible. Read/write  **Boolean**.


## Syntax

_expression_. `Visible`

_expression_ A variable that represents a [NavigationControl](Access.NavigationControl.md) object.


## Remarks

To hide an object when printing, use the  **DisplayWhen** property.

You can use the  **Visible** property to hide a control on a form or report by including the property in a macro or event procedure that runs when the **Current** event occurs. For example, you can show or hide a congratulatory message next to a salesperson's monthly sales total in a sales report, depending on the sales total.


## See also


[NavigationControl Object](Access.NavigationControl.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]