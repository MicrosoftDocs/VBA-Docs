---
title: Section.Height property (Access)
keywords: vbaac10.chm12200
f1_keywords:
- vbaac10.chm12200
ms.prod: access
api_name:
- Access.Section.Height
ms.assetid: 7e568d9f-518b-6d26-e960-dac84e93d45b
ms.date: 02/21/2019
localization_priority: Normal
---


# Section.Height property (Access)

Gets or sets the height of the specified object in [twips](../language/glossary/vbe-glossary.md#twip). Read/write **Integer**.


## Syntax

_expression_.**Height**

_expression_ A variable that represents a **[Section](Access.Section.md)** object.

## Remarks

The **Height** property applies only to form sections and report sections, not to forms and reports.

For report sections, you can't use a macro or Visual Basic to set the **Height** property when you print or preview a report. For report controls, you can set the **Height** property when you print or preview a report only by using a macro or an event procedure specified in a section's **OnFormat** event property setting.

Microsoft Access automatically sets the **Height** property when you create or size a control or when you size a window in form Design view or report Design view.

The height of sections is measured from the inside of their borders. The height of controls is measured from the center of their borders so that controls with different border widths align correctly. 




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]

