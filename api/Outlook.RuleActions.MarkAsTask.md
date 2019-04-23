---
title: RuleActions.MarkAsTask property (Outlook)
keywords: vbaol11.chm2198
f1_keywords:
- vbaol11.chm2198
ms.prod: outlook
api_name:
- Outlook.RuleActions.MarkAsTask
ms.assetid: 9dd48e1a-d780-0923-11b0-e980c1fe19ab
ms.date: 06/08/2017
localization_priority: Normal
---


# RuleActions.MarkAsTask property (Outlook)

Returns a  **[MarkAsTaskRuleAction](Outlook.MarkAsTaskRuleAction.md)** object with **[MarkAsTaskRuleAction.ActionType](Outlook.MarkAsTaskRuleAction.ActionType.md)** being **olRuleActionMarkAsTask**. Read-only.


## Syntax

_expression_. `MarkAsTask`

_expression_ A variable that represents a [RuleActions](Outlook.RuleActions.md) object.


## Remarks

Use the returned  **MarkAsTaskRuleAction** object when enumerating the rule actions of an existing rule or when creating a new rule that specifies the action of marking a message as a task.

This property of the  **[RuleActions](Outlook.RuleActions.md)** collection always returns a **MarkAsTaskRuleAction** object regardless of whether the rule associated with this **RuleActions** collection has defined such a rule action. If the rule has defined and enabled such a rule action, then **[MarkAsTaskRuleAction.Enabled](Outlook.MarkAsTaskRuleAction.Enabled.md)** will be **True**.


## See also


[RuleActions Object](Outlook.RuleActions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]