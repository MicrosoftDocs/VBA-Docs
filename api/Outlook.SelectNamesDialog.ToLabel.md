---
title: SelectNamesDialog.ToLabel property (Outlook)
keywords: vbaol11.chm830
f1_keywords:
- vbaol11.chm830
ms.prod: outlook
api_name:
- Outlook.SelectNamesDialog.ToLabel
ms.assetid: 1c2f15fd-57c6-e0a5-923c-2b3b217bb7a0
ms.date: 06/08/2017
localization_priority: Normal
---


# SelectNamesDialog.ToLabel property (Outlook)

Returns or sets a **String** for the text that appears on the **To** command button on the **Select Names** dialog box. Read/write.

## Syntax

_expression_. `ToLabel`

_expression_ A variable that represents a [SelectNamesDialog](Outlook.SelectNamesDialog.md) object.

## Remarks

To provide an accelerator key for the recipient edit boxes, include an ampersand (&) character in the label argument string, immediately before the character that serves as the access key. For example, if **ToLabel** is the string "Local & Attendees", users can press **ALT+A** to move the focus to the first recipient edit box.

If **ToLabel** is not set, the default value will be the localized string for "To". If the **ToLabel** is set to an empty string, the **To** command button shows **-&gt;**. If the **ToLabel** property contains more than 32 characters (including the ampersand (&) access key), only the first 32 characters will be displayed in the command button.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]