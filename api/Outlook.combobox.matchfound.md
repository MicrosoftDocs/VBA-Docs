---
title: ComboBox.MatchFound Property (Outlook Forms Script)
keywords: olfm10.chm2001490
f1_keywords:
- olfm10.chm2001490
ms.prod: outlook
ms.assetid: 2e35541f-990d-fa2a-4431-695f9d951c98
ms.date: 06/08/2017
localization_priority: Normal
---


# ComboBox.MatchFound Property (Outlook Forms Script)

Returns a **Boolean** value that indicates whether the text that a user has typed into a **[ComboBox](Outlook.combobox.md)** matches any of the entries in the list. Read-only.


## Syntax

_expression_.**MatchFound**

_expression_ A variable that represents a **ComboBox** object.


## Remarks

 **True** if the contents of the **[Value](Outlook.combobox.value.md)** property matches one of the records in the list. **False** if the contents of **Value** does not match any of the records in the list (default).

The **MatchFound** property is read-only. It is not applicable when the **[MatchEntry](Outlook.combobox.matchentry.md)** property is set to 2.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]