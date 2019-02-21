---
title: Page.Height property (Access)
keywords: vbaac10.chm12158
f1_keywords:
- vbaac10.chm12158
ms.prod: access
api_name:
- Access.Page.Height
ms.assetid: df6c7cc3-bcf5-6607-144a-383a1f26d21e
ms.date: 02/21/2019
localization_priority: Normal
---


# Page.Height property (Access)

Gets or sets the height of the specified object in [twips](../language/glossary/vbe-glossary.md#twip). Read/write **Integer**.


## Syntax

_expression_.**Height**

_expression_ A variable that represents a **[Page](Access.Page.md)** object.


## Remarks

For report controls, you can set the **Height** property when you print or preview a report only by using a macro or an event procedure specified in a section's **OnFormat** event property setting.

Microsoft Access automatically sets the **Height** property when you create or size a control or when you size a window in form Design view or report Design view.

The height of controls is measured from the center of their borders so that controls with different border widths align correctly. 




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]