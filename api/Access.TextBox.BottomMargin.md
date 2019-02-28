---
title: TextBox.BottomMargin property (Access)
keywords: vbaac10.chm11142
f1_keywords:
- vbaac10.chm11142
ms.prod: access
api_name:
- Access.TextBox.BottomMargin
ms.assetid: a6ef1155-24c8-1254-614b-c912fda8dae2
ms.date: 02/28/2019
localization_priority: Normal
---


# TextBox.BottomMargin property (Access)

Along with the **LeftMargin**, **RightMargin**, and **TopMargin** properties, specifies the location of information displayed within a text box control. Read/write **Integer**.


## Syntax

_expression_.**BottomMargin**

_expression_ A variable that represents a **[TextBox](Access.TextBox.md)** object.


## Remarks

A control's displayed information location is the distance measured from the control's left, top, right, or bottom border to the left, top, right, or bottom edge of the displayed information. To use a unit of measurement different from the setting in the regional settings of Windows, specify the unit (for example, cm or in).

In Visual Basic, use a numeric expression to set the value of this property. Values are expressed in [twips](../language/glossary/vbe-glossary.md#twip).


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]