---
title: RuleActions.MoveToFolder property (Outlook)
keywords: vbaol11.chm2191
f1_keywords:
- vbaol11.chm2191
ms.prod: outlook
api_name:
- Outlook.RuleActions.MoveToFolder
ms.assetid: 6d9c577d-e022-72fc-45f2-bdda7a8761de
ms.date: 06/08/2017
localization_priority: Normal
---


# RuleActions.MoveToFolder property (Outlook)

Returns a **[MoveOrCopyRuleAction](Outlook.MoveOrCopyRuleAction.md)** object with **[MoveOrCopyRuleAction.ActionType](Outlook.MoveOrCopyRuleAction.ActionType.md)** being **olRuleActionMoveToFolder**. Read-only.


## Syntax

_expression_. `MoveToFolder`

_expression_ A variable that represents a [RuleActions](Outlook.RuleActions.md) object.


## Remarks

Use the returned  **MoveOrCopyRuleAction** object when enumerating the rule actions of an existing rule or when creating a new rule that specifies copying a message to a specific folder as an action.

This property of the  **[RuleActions](Outlook.RuleActions.md)** collection always returns a **MoveOrCopyRuleAction** object regardless of whether the rule associated with this **RuleActions** collection has defined such a rule action. If the rule has defined and enabled such a rule action, then **[MoveOrCopyRuleAction.Enabled](Outlook.MoveOrCopyRuleAction.Enabled.md)** will be **True**.


## See also


[RuleActions Object](Outlook.RuleActions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]