---
title: OlkTextBox.SelStart property (Outlook)
keywords: vbaol11.chm1000062
f1_keywords:
- vbaol11.chm1000062
ms.prod: outlook
api_name:
- Outlook.OlkTextBox.SelStart
ms.assetid: cca8ffc2-4c68-72f5-7e09-6f8845d72e35
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkTextBox.SelStart property (Outlook)

Returns or sets a  **Long** that specifies either the starting point of the selected text or the insertion point if no text has been selected. Read/write.


## Syntax

_expression_. `SelStart`

_expression_ A variable that represents an [OlkTextBox](Outlook.OlkTextBox.md) object.


## Remarks

The current selection is specified by  **[SelText](Outlook.OlkTextBox.SelText.md)**, which is a portion of the control's **[Value](Outlook.OlkTextBox.Value.md)**. The maximum number of characters that can be supported for **Value** is **[MaxLength](Outlook.OlkTextBox.MaxLength.md)**.

The default value is zero, which means no text is selected and the insertion point is at the beginning.

The  **SelStart** property is always valid, even when the control does not have focus. Setting **SelStart** to a value less than zero causes an error. Setting **SelStart** to a value greater than **MaxLength** will reset **SelStart** to **MaxLength**. Changing the value of **SelStart** cancels any existing selection, places the insertion point in the text, and sets the **SelLength** property to zero.


## See also


[OlkTextBox Object](Outlook.OlkTextBox.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]