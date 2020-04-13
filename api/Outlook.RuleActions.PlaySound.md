---
title: RuleActions.PlaySound property (Outlook)
keywords: vbaol11.chm2197
f1_keywords:
- vbaol11.chm2197
ms.prod: outlook
api_name:
- Outlook.RuleActions.PlaySound
ms.assetid: 43a79f2d-9e7b-7053-6901-40e815220ac0
ms.date: 06/08/2017
localization_priority: Normal
---


# RuleActions.PlaySound property (Outlook)

Returns a **[PlaySoundRuleAction](Outlook.PlaySoundRuleAction.md)** object with **[PlaySoundRuleAction.ActionType](Outlook.PlaySoundRuleAction.ActionType.md)** being **olRuleActionNotifyRead**. Read-only.


## Syntax

_expression_. `PlaySound`

_expression_ A variable that represents a [RuleActions](Outlook.RuleActions.md) object.


## Remarks

Use the returned  **PlaySoundRuleAction** object when enumerating the rule actions of an existing rule or when creating a new rule that specifies playing a sound file as an action.

This property of the  **[RuleActions](Outlook.RuleActions.md)** collection always returns a **PlaySoundRuleAction** object regardless of whether the rule associated with this **RuleActions** collection has defined such a rule action. If the rule has defined and enabled such a rule action, then **[PlaySoundRuleAction.Enabled](Outlook.PlaySoundRuleAction.Enabled.md)** will be **True**.


## See also


[RuleActions Object](Outlook.RuleActions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]