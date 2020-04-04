---
title: OlkComboBox.TopIndex property (Outlook)
keywords: vbaol11.chm1000217
f1_keywords:
- vbaol11.chm1000217
ms.prod: outlook
api_name:
- Outlook.OlkComboBox.TopIndex
ms.assetid: 483db226-bf25-55e6-d453-a494747ff7d9
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkComboBox.TopIndex property (Outlook)

Returns or sets a **Long** that represents the index of the item at the top of the displayed portion of the list in the combo box. Read/write.


## Syntax

_expression_. `TopIndex`

_expression_ A variable that represents an [OlkComboBox](Outlook.OlkComboBox.md) object.


## Remarks

As the list scrolls, the item at the top of the list will change, and the value of this property will change to reflect the item currently displayed at the top of the list.

The index value is zero-based. The default value is -1, indicating that no special ordering should be applied.


## See also


[OlkComboBox Object](Outlook.OlkComboBox.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]