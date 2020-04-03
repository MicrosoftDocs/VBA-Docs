---
title: OlkComboBox.SelText property (Outlook)
keywords: vbaol11.chm1000223
f1_keywords:
- vbaol11.chm1000223
ms.prod: outlook
api_name:
- Outlook.OlkComboBox.SelText
ms.assetid: 595b3e85-7d30-72bc-c1d4-b45c4492c221
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkComboBox.SelText property (Outlook)

Returns a **String** that represents the selected portion of the value of the combo box. Read-only.


## Syntax

_expression_. `SelText`

_expression_ A variable that represents an [OlkComboBox](Outlook.OlkComboBox.md) object.


## Remarks

 **SelText** represents the current selection, which is a portion of the control's **[Value](Outlook.OlkComboBox.Value.md)**. The maximum number of characters that can be supported for **Value** is **[MaxLength](Outlook.OlkComboBox.MaxLength.md)**.

The default value is an empty string.


## See also


[OlkComboBox Object](Outlook.OlkComboBox.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]