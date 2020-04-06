---
title: Rule.Actions property (Outlook)
keywords: vbaol11.chm2174
f1_keywords:
- vbaol11.chm2174
ms.prod: outlook
api_name:
- Outlook.Rule.Actions
ms.assetid: 2b1e2ad4-c735-b3a8-6b27-5004f10393ce
ms.date: 06/08/2017
localization_priority: Normal
---


# Rule.Actions property (Outlook)

Returns a **[RuleActions](Outlook.RuleActions.md)** collection object that represents all the available rule actions for the rule. Read-only.


## Syntax

_expression_. `Actions`

_expression_ A variable that represents a [Rule](Outlook.Rule.md) object.


## Remarks

You can enumerate and enable rules with any rule action that the Rules and Alerts Wizard support, but you can programmatically create rules that have only the most commonly used rule actions, and not any rule action that the Rules and Alerts Wizard supports. For more information on rule action support, see [Specify Rule Actions](../outlook/How-to/Rules/specifying-rule-actions.md).

Through the  **Actions** property, each rule is associated with a **RuleActions** object. The **RuleActions** collection is a fixed object - you cannot add or remove items from this collection. Rule actions that are enabled in the rule will have an enabled rule action in the **RuleActions** collection. Rule actions that are not enabled in the rule will have a rule action in this collection that has the **[RuleAction.Enabled](Outlook.RuleAction.Enabled.md)** property set to **False**. Rule actions that are not supported during programmatic rule creation can only be enumerated in the **RuleActions** collection for an existing rule, but because the **RuleActions** collection is fixed, you cannot create a rule and add such an action to the associated **RuleActions** collection.


## See also


[Rule Object](Outlook.Rule.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]