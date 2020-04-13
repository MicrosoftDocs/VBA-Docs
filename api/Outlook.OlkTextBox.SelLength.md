---
title: OlkTextBox.SelLength property (Outlook)
keywords: vbaol11.chm1000063
f1_keywords:
- vbaol11.chm1000063
ms.prod: outlook
api_name:
- Outlook.OlkTextBox.SelLength
ms.assetid: 89d040ba-b28f-20f1-e449-1c533370b711
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkTextBox.SelLength property (Outlook)

Returns or sets a **Long** that specifies the number of characters in the current selection. Read/write.


## Syntax

_expression_. `SelLength`

_expression_ A variable that represents an [OlkTextBox](Outlook.OlkTextBox.md) object.


## Remarks

The current selection is specified by  **[SelText](Outlook.OlkTextBox.SelText.md)**, which is a portion of the control's **[Value](Outlook.OlkTextBox.Value.md)**. The maximum number of characters that can be supported for **Value** is **[MaxLength](Outlook.OlkTextBox.MaxLength.md)**.

The default value is zero, which means no text is currently selected.

The **SelLength** property is always valid, even when the control does not have focus.

Setting  **SelLength** to a value less than zero causes an error. Attempting to set the value greater than **MaxLength** results in setting **SelLength** to **MaxLength**.


## See also


[OlkTextBox Object](Outlook.OlkTextBox.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]