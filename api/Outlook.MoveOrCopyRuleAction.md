---
title: MoveOrCopyRuleAction object (Outlook)
keywords: vbaol11.chm3164
f1_keywords:
- vbaol11.chm3164
ms.prod: outlook
api_name:
- Outlook.MoveOrCopyRuleAction
ms.assetid: db951ad8-0d05-1696-acf4-c1da4fbdee33
ms.date: 06/08/2017
localization_priority: Normal
---


# MoveOrCopyRuleAction object (Outlook)

Represents an action that moves or copies a message.


## Remarks

 **MoveOrCopyRuleAction** is derived from the **[RuleAction](Outlook.RuleAction.md)** object. Each rule is associated with a **[RuleActions](Outlook.RuleActions.md)** object which has a **[CopyToFolder](Outlook.RuleActions.CopyToFolder.md)** property and a **[MoveToFolder](Outlook.RuleActions.MoveToFolder.md)** property. Each of these properties always returns a **MoveOrCopyRuleAction** object. **[MoveOrCopyRuleAction.ActionType](Outlook.MoveOrCopyRuleAction.ActionType.md)** distinguishes between the two action types. If the rule has an enabled rule action that copies or moves a message to the specified folder, then the corresponding **[MoveOrCopyRuleAction.Enabled](Outlook.MoveOrCopyRuleAction.Enabled.md)** would be **True**.

For more information on specifying rule actions, see [Specify Rule Actions](../outlook/How-to/Rules/specifying-rule-actions.md).


## Properties



|Name|
|:-----|
|[ActionType](Outlook.MoveOrCopyRuleAction.ActionType.md)|
|[Application](Outlook.MoveOrCopyRuleAction.Application.md)|
|[Class](Outlook.MoveOrCopyRuleAction.Class.md)|
|[Enabled](Outlook.MoveOrCopyRuleAction.Enabled.md)|
|[Folder](Outlook.MoveOrCopyRuleAction.Folder.md)|
|[Parent](Outlook.MoveOrCopyRuleAction.Parent.md)|
|[Session](Outlook.MoveOrCopyRuleAction.Session.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]