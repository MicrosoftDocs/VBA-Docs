---
title: NavigationControl.Height property (Access)
keywords: vbaac10.chm11074
f1_keywords:
- vbaac10.chm11074
ms.prod: access
api_name:
- Access.NavigationControl.Height
ms.assetid: bf46c094-6eef-452b-dca9-ff6d4a3e5006
ms.date: 02/21/2019
localization_priority: Normal
---


# NavigationControl.Height property (Access)

Gets or sets the height of the specified object in [twips](../language/glossary/vbe-glossary.md#twip). Read/write **Integer**.


## Syntax

_expression_.**Height**

_expression_ A variable that represents a **[NavigationControl](Access.NavigationControl.md)** object.


## Remarks

For report controls, you can set the **Height** property when you print or preview a report only by using a macro or an event procedure specified in a section's **OnFormat** event property setting.

Microsoft Access automatically sets the **Height** property when you create or size a control or when you size a window in form Design view or report Design view.

The height of controls is measured from the center of their borders so that controls with different border widths align correctly. 




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]