---
title: RuleActions.DesktopAlert property (Outlook)
keywords: vbaol11.chm2187
f1_keywords:
- vbaol11.chm2187
ms.prod: outlook
api_name:
- Outlook.RuleActions.DesktopAlert
ms.assetid: 700c3e5a-ebb1-3cfe-e27d-eea305c27143
ms.date: 06/08/2017
localization_priority: Normal
---


# RuleActions.DesktopAlert property (Outlook)

Returns a  **[RuleAction](Outlook.RuleAction.md)** object with **[RuleAction.ActionType](Outlook.RuleAction.ActionType.md)** being **olRuleActionDesktopAlert**. Read-only.


## Syntax

_expression_. `DesktopAlert`

_expression_ A variable that represents a [RuleActions](Outlook.RuleActions.md) object.


## Remarks

Use the returned  **RuleAction** object when enumerating the rule actions of an existing rule or when creating a new rule that specifies displaying a desktop alert as an action.

This property of the  **[RuleActions](Outlook.RuleActions.md)** collection always returns a **RuleAction** object regardless of whether the rule associated with this **RuleActions** collection has defined such a rule action. If the rule has defined and enabled such a rule action, then **[RuleAction.Enabled](Outlook.MoveOrCopyRuleAction.Enabled.md)** will be **True**.


## See also


[RuleActions Object](Outlook.RuleActions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]