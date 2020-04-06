---
title: Rule.Exceptions property (Outlook)
keywords: vbaol11.chm2176
f1_keywords:
- vbaol11.chm2176
ms.prod: outlook
api_name:
- Outlook.Rule.Exceptions
ms.assetid: 843c2690-ee39-bac7-d593-80c3dd31087f
ms.date: 06/08/2017
localization_priority: Normal
---


# Rule.Exceptions property (Outlook)

Returns a **[RuleConditions](Outlook.RuleConditions.md)** collection object that represents all the available rule exception conditions for the rule. Read-only.


## Syntax

_expression_. `Exceptions`

_expression_ A variable that represents a [Rule](Outlook.Rule.md) object.


## Remarks

An exception condition for a rule states the condition under which the rule should not be applied. Both the  **[Conditions](Outlook.Rule.Conditions.md)** and **Exceptions** properties share the same pool of conditions and return a corresponding **RuleConditions** collection object.

You can enumerate and enable rules with any rule exception condition that the Rules and Alerts Wizard support, but you can programmatically create rules that have only the most commonly used rule exception conditions, and not any rule exception condition that the Rules and Alerts Wizard supports. For more information on rule condition support, see [Specify Rule Conditions](../outlook/How-to/Rules/specifying-rule-conditions.md).

Through the  **Conditions** property, each rule is associated with a **RuleConditions** object. The **RuleConditions** collection is a fixed object - you cannot add or remove items from this collection. Rule exception conditions that are enabled in the rule will have an enabled rule exception condition in the **RuleConditions** collection. Rule exception conditions that are not enabled in the rule will have a rule exception condition in this collection that has the **[RuleCondition.Enabled](Outlook.RuleCondition.Enabled.md)** property set to **False**. Rule exception conditions that are not supported during programmatic rule creation can only be enumerated in the **RuleConditions** collection for an existing rule, but because the **RuleConditions** collection is fixed, you cannot create a rule and add such an exception condition to the associated **RuleConditions** collection.


## See also


[Rule Object](Outlook.Rule.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]