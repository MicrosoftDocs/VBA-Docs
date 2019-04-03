---
title: MarkAsTaskRuleAction object (Outlook)
keywords: vbaol11.chm3170
f1_keywords:
- vbaol11.chm3170
ms.prod: outlook
api_name:
- Outlook.MarkAsTaskRuleAction
ms.assetid: 639d9242-7387-2b25-9d0f-f7a14cf16790
ms.date: 06/08/2017
localization_priority: Normal
---


# MarkAsTaskRuleAction object (Outlook)

Represents an action that marks a message as a task.


## Remarks

 **MarkAsTaskRuleAction** is derived from the **[RuleAction](Outlook.RuleAction.md)** object. Each rule is associated with a **[RuleActions](Outlook.RuleActions.md)** object which has a **[MarkAsTask](Outlook.RuleActions.MarkAsTask.md)** property. The **MarkAsTask** property always returns a **MarkAsTaskRuleAction** object. If the rule has an enabled rule action that marks a message as a task, then **[MarkAsTaskRuleAction.Enabled](Outlook.MarkAsTaskRuleAction.Enabled.md)** would be **True**.

For more information on specifying rule actions, see [Specify Rule Actions](../outlook/How-to/Rules/specifying-rule-actions.md).


## Properties



|Name|
|:-----|
|[ActionType](Outlook.MarkAsTaskRuleAction.ActionType.md)|
|[Application](Outlook.MarkAsTaskRuleAction.Application.md)|
|[Class](Outlook.MarkAsTaskRuleAction.Class.md)|
|[Enabled](Outlook.MarkAsTaskRuleAction.Enabled.md)|
|[FlagTo](Outlook.MarkAsTaskRuleAction.FlagTo.md)|
|[MarkInterval](Outlook.MarkAsTaskRuleAction.MarkInterval.md)|
|[Parent](Outlook.MarkAsTaskRuleAction.Parent.md)|
|[Session](Outlook.MarkAsTaskRuleAction.Session.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]