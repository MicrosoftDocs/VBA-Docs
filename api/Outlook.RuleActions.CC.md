---
title: RuleActions.CC property (Outlook)
keywords: vbaol11.chm2192
f1_keywords:
- vbaol11.chm2192
ms.prod: outlook
api_name:
- Outlook.RuleActions.CC
ms.assetid: edbaaf74-cfd2-304b-61f3-8d12a621239c
ms.date: 06/08/2017
localization_priority: Normal
---


# RuleActions.CC property (Outlook)

Returns a  **[SendRuleAction](Outlook.SendRuleAction.md)** object with **[SendRuleAction.ActionType](Outlook.SendRuleAction.ActionType.md)** being **olRuleActionCcMessage**. Read-only.


## Syntax

_expression_. `CC`

_expression_ A variable that represents a [RuleActions](Outlook.RuleActions.md) object.


## Remarks

Use the returned  **SendRuleAction** object when enumerating the rule actions of an existing rule or when creating a new rule that specifies cc-ing a message to specific recipients as an action.

This property of the  **[RuleActions](Outlook.RuleActions.md)** collection always returns a **SendRuleAction** object regardless of whether the rule associated with this **RuleActions** collection has defined such a rule action. If the rule has defined and enabled such a rule action, then **[SendRuleAction.Enabled](Outlook.SendRuleAction.Enabled.md)** will be **True**.


## See also


[RuleActions Object](Outlook.RuleActions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]