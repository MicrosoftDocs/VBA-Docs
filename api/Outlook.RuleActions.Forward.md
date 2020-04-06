---
title: RuleActions.Forward property (Outlook)
keywords: vbaol11.chm2193
f1_keywords:
- vbaol11.chm2193
ms.prod: outlook
api_name:
- Outlook.RuleActions.Forward
ms.assetid: 48315808-5ef7-3262-a305-5b659986e7a8
ms.date: 06/08/2017
localization_priority: Normal
---


# RuleActions.Forward property (Outlook)

Returns a **[SendRuleAction](Outlook.SendRuleAction.md)** object with **[SendRuleAction.ActionType](Outlook.SendRuleAction.ActionType.md)** being **olRuleActionForward**. Read-only.


## Syntax

_expression_. `Forward`

_expression_ A variable that represents a [RuleActions](Outlook.RuleActions.md) object.


## Remarks

Use the returned  **SendRuleAction** object when enumerating the rule actions of an existing rule or when creating a new rule that specifies forwarding a message to specific recipients as an action.

This property of the  **[RuleActions](Outlook.RuleActions.md)** collection always returns a **SendRuleAction** object regardless of whether the rule associated with this **RuleActions** collection has defined such a rule action. If the rule has defined and enabled such a rule action, then **[SendRuleAction.Enabled](Outlook.SendRuleAction.Enabled.md)** will be **True**.


## See also


[RuleActions Object](Outlook.RuleActions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]