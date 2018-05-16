---
title: PageBreak.Visible Property (Access)
keywords: vbaac10.chm11678
f1_keywords:
- vbaac10.chm11678
ms.prod: access
api_name:
- Access.PageBreak.Visible
ms.assetid: bce10ac3-a7a5-5d0e-df76-b8222aa64267
ms.date: 06/08/2017
---


# PageBreak.Visible Property (Access)

Returns or sets whether the object is visible. Read/write  **Boolean**.


## Syntax

 _expression_. **Visible**

 _expression_ A variable that represents a **PageBreak** object.


## Remarks

To hide an object when printing, use the  **DisplayWhen** property.

You can use the  **Visible** property to hide a control on a form or report by including the property in a macro or event procedure that runs when the **Current** event occurs. For example, you can show or hide a congratulatory message next to a salesperson's monthly sales total in a sales report, depending on the sales total.


## See also


#### Concepts


[PageBreak Object](Access.PageBreak.md)

