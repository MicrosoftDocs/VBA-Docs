---
title: RuleActions.AssignToCategory property (Outlook)
keywords: vbaol11.chm2196
f1_keywords:
- vbaol11.chm2196
ms.prod: outlook
api_name:
- Outlook.RuleActions.AssignToCategory
ms.assetid: 7780487b-3dd4-6143-2250-2109872b6192
ms.date: 06/08/2017
localization_priority: Normal
---


# RuleActions.AssignToCategory property (Outlook)

Returns an  **[AssignToCategoryRuleAction](Outlook.AssignToCategoryRuleAction.md)** object with **[AssignToCategoryRuleAction.ActionType](Outlook.AssignToCategoryRuleAction.ActionType.md)** being **olRuleAssignToCategory**. Read-only.


## Syntax

_expression_. `AssignToCategory`

_expression_ A variable that represents a [RuleActions](Outlook.RuleActions.md) object.


## Remarks

Use the returned  **AssignToCategoryRuleAction** object when enumerating the rule actions of an existing rule or when creating a new rule that assigns categories to a message.

This property of the  **[RuleActions](Outlook.RuleActions.md)** collection always returns an **AssignToCategoryRuleAction** object regardless of whether the rule associated with this **RuleActions** collection has defined such a rule action. If the rule has defined and enabled such a rule action, then **[AssignToCategoryRuleAction.Enabled](Outlook.AssignToCategoryRuleAction.Enabled.md)** will be **True**.


## See also


[RuleActions Object](Outlook.RuleActions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]