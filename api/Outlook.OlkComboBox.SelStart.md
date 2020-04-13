---
title: OlkComboBox.SelStart property (Outlook)
keywords: vbaol11.chm1000221
f1_keywords:
- vbaol11.chm1000221
ms.prod: outlook
api_name:
- Outlook.OlkComboBox.SelStart
ms.assetid: f3141a7c-b9a5-b738-8803-9100e2283dc1
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkComboBox.SelStart property (Outlook)

Returns or sets a **Long** that specifies either the starting point of the selected text or the insertion point if no text has been selected. Read/write.


## Syntax

_expression_. `SelStart`

_expression_ A variable that represents an [OlkComboBox](Outlook.OlkComboBox.md) object.


## Remarks

The current selection is specified by  **[SelText](Outlook.OlkComboBox.SelText.md)**, which is a portion of the control's **[Value](Outlook.OlkComboBox.Value.md)**. The maximum number of characters that can be supported for **Value** is **[MaxLength](Outlook.OlkComboBox.MaxLength.md)**.

The default value is zero, which means no text is selected and the insertion point is at the beginning.

The **SelStart** property is always valid, even when the control does not have focus. Setting **SelStart** to a value less than zero causes an error. Setting **SelStart** to a value greater than **MaxLength** will reset **SelStart** to **MaxLength**. Changing the value of **SelStart** cancels any existing selection, places the insertion point in the text, and sets the **SelLength** property to zero.


## See also


[OlkComboBox Object](Outlook.OlkComboBox.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]