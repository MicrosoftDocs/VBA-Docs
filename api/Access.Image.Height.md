---
title: Image.Height property (Access)
keywords: vbaac10.chm10379
f1_keywords:
- vbaac10.chm10379
ms.prod: access
api_name:
- Access.Image.Height
ms.assetid: cfc48a43-736f-58c0-95a5-fe248de8b9d3
ms.date: 02/21/2019
localization_priority: Normal
---


# Image.Height property (Access)

Gets or sets the height of the specified object in [twips](../language/glossary/vbe-glossary.md#twip). Read/write **Integer**.


## Syntax

_expression_.**Height**

_expression_ A variable that represents an **[Image](Access.Image.md)** object.


## Return value

Integer


## Remarks

For report controls, you can set the **Height** property when you print or preview a report only by using a macro or an event procedure specified in a section's **OnFormat** event property setting.

Microsoft Access automatically sets the **Height** property when you create or size a control or when you size a window in form Design view or report Design view.

The height of controls is measured from the center of their borders so that controls with different border widths align correctly. 




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]