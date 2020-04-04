---
title: RuleActions.NotifyRead property (Outlook)
keywords: vbaol11.chm2189
f1_keywords:
- vbaol11.chm2189
ms.prod: outlook
api_name:
- Outlook.RuleActions.NotifyRead
ms.assetid: 922a1ea7-8992-0387-e4e1-2e74d6a2cf2a
ms.date: 06/08/2017
localization_priority: Normal
---


# RuleActions.NotifyRead property (Outlook)

Returns a **[RuleAction](Outlook.RuleAction.md)** object with **[RuleAction.ActionType](Outlook.RuleAction.ActionType.md)** being **olRuleActionNotifyRead**. Read-only.


## Syntax

_expression_. `NotifyRead`

_expression_ A variable that represents a [RuleActions](Outlook.RuleActions.md) object.


## Remarks

Use the returned  **RuleAction** object when enumerating the rule actions of an existing rule or when creating a new rule that specifies sending a notification about the opening of a message as an action.

This property of the  **[RuleActions](Outlook.RuleActions.md)** collection always returns a **RuleAction** object regardless of whether the rule associated with this **RuleActions** collection has defined such a rule action. If the rule has defined and enabled such a rule action, then **[RuleAction.Enabled](Outlook.MoveOrCopyRuleAction.Enabled.md)** will be **True**.


## See also


[RuleActions Object](Outlook.RuleActions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]