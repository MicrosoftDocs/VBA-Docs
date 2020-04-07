---
title: MoveOrCopyRuleAction.Folder property (Outlook)
keywords: vbaol11.chm2214
f1_keywords:
- vbaol11.chm2214
ms.prod: outlook
api_name:
- Outlook.MoveOrCopyRuleAction.Folder
ms.assetid: d9588bab-c863-428f-0bbe-289f0981ff0f
ms.date: 06/08/2017
localization_priority: Normal
---


# MoveOrCopyRuleAction.Folder property (Outlook)

Returns or sets a **[Folder](Outlook.Folder.md)** object that represents the folder to which the rule moves or copies the message. Read/write.


## Syntax

_expression_. `Folder`

_expression_ A variable that represents a [MoveOrCopyRuleAction](Outlook.MoveOrCopyRuleAction.md) object.


## Remarks

If no folder has been assigned to the move or copy rule action, this property is **Null** (**Nothing** in Visual Basic).

This property returns an error if the specified folder cannot serve as a target folder for the copy or move operation. For example, the folder is a search folder, is read-only, or the user does not have the required permissions to move or copy messages to it.


## See also


[MoveOrCopyRuleAction Object](Outlook.MoveOrCopyRuleAction.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]