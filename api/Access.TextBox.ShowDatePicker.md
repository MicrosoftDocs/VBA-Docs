---
title: TextBox.ShowDatePicker property (Access)
keywords: vbaac10.chm14293
f1_keywords:
- vbaac10.chm14293
ms.prod: access
api_name:
- Access.TextBox.ShowDatePicker
ms.assetid: 5d65938b-ac7b-abbd-2e50-41f41c0b1558
ms.date: 03/26/2019
localization_priority: Normal
---


# TextBox.ShowDatePicker property (Access)

Gets or sets whether the date picker control is displayed for the specified text box. Read/write **Integer**.


## Syntax

_expression_.**ShowDatePicker**

_expression_ A variable that represents a **[TextBox](Access.TextBox.md)** object.


## Remarks

The **ShowDatePicker** property uses the following settings.

|Value|Description|
|:-----|:-----|
|0|The date picker control is not displayed.|
|1|The date picker control is displayed when the text box is bound to a **Date** field.|

The **ShowDatePicker** property is not available for text boxes on reports.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
