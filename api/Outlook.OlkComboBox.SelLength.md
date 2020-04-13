---
title: OlkComboBox.SelLength property (Outlook)
keywords: vbaol11.chm1000222
f1_keywords:
- vbaol11.chm1000222
ms.prod: outlook
api_name:
- Outlook.OlkComboBox.SelLength
ms.assetid: 3cbd5016-3868-6cf9-c28c-8d692620f367
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkComboBox.SelLength property (Outlook)

Returns or sets a **Long** that specifies the number of characters in the current selection. Read/write.


## Syntax

_expression_. `SelLength`

_expression_ A variable that represents an [OlkComboBox](Outlook.OlkComboBox.md) object.


## Remarks

The current selection is specified by  **[SelText](Outlook.OlkComboBox.SelText.md)**, which is a portion of the control's **[Value](Outlook.OlkComboBox.Value.md)**. The maximum number of characters that can be supported for **Value** is **[MaxLength](Outlook.OlkComboBox.MaxLength.md)**.

The default value is zero, which means no text is currently selected.

The **SelLength** property is always valid, even when the control does not have focus.

Setting  **SelLength** to a value less than zero causes an error. Attempting to set the value greater than **MaxLength** results in setting **SelLength** to **MaxLength**.


## See also


[OlkComboBox Object](Outlook.OlkComboBox.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]