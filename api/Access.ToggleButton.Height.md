---
title: ToggleButton.Height property (Access)
keywords: vbaac10.chm11721
f1_keywords:
- vbaac10.chm11721
ms.prod: access
api_name:
- Access.ToggleButton.Height
ms.assetid: 8544f955-3891-3799-5207-de7fa2a5a224
ms.date: 02/21/2019
localization_priority: Normal
---


# ToggleButton.Height property (Access)

Gets or sets the height of the specified object in [twips](../language/glossary/vbe-glossary.md#twip). Read/write **Integer**.


## Syntax

_expression_.**Height**

_expression_ A variable that represents a **[ToggleButton](Access.ToggleButton.md)** object.


## Remarks

For report controls, you can set the **Height** property when you print or preview a report only by using a macro or an event procedure specified in a section's **OnFormat** event property setting.

Microsoft Access automatically sets the **Height** property when you create or size a control or when you size a window in form Design view or report Design view.

The height of controls is measured from the center of their borders so that controls with different border widths align correctly. 




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]