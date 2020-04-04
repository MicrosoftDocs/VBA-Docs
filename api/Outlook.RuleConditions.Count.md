---
title: RuleConditions.Count property (Outlook)
keywords: vbaol11.chm2300
f1_keywords:
- vbaol11.chm2300
ms.prod: outlook
api_name:
- Outlook.RuleConditions.Count
ms.assetid: 7950c105-4528-40aa-f263-b800a68ae1ad
ms.date: 06/08/2017
localization_priority: Normal
---


# RuleConditions.Count property (Outlook)

Returns a **Long** indicating the count of objects in the specified collection. Read-only.


## Syntax

_expression_.**Count**

_expression_ A variable that represents a [RuleConditions](Outlook.RuleConditions.md) object.


## Remarks

You can enumerate the  **[RuleConditions](Outlook.RuleConditions.md)** collection from 1 through **RuleConditions.Count** to determine all the rule conditions for an existing rule. Although the Rules OM supports creating rules with only the most commonly used rule conditions and not all rule conditions supported by the Rules and Alerts Wizard, the **RuleConditions** collection includes all rule conditions of a rule. Hence you can always enumerate the **RuleConditions** collection object to determine which rule conditions are enabled for an existing rule.


## See also


[RuleConditions Object](Outlook.RuleConditions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]