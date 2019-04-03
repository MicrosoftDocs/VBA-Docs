---
title: FormNameRuleCondition object (Outlook)
keywords: vbaol11.chm3180
f1_keywords:
- vbaol11.chm3180
ms.prod: outlook
api_name:
- Outlook.FormNameRuleCondition
ms.assetid: 75b7f687-66e6-4863-b8aa-f19e98fedc45
ms.date: 06/08/2017
localization_priority: Normal
---


# FormNameRuleCondition object (Outlook)

Represents a rule condition that evaluates whether a form name was used to send or receive an item.


## Remarks

 **FormNameRuleCondition** is derived from the **[RuleCondition](Outlook.RuleCondition.md)** object. Each rule is associated with a **[RuleConditions](Outlook.RuleConditions.md)** object which has a **[FormName](Outlook.RuleConditions.FormName.md)** property. The **FormName** property always returns a **FormNameRuleCondition** object. If the rule has any of these rule conditions enabled, then **[FormNameRuleCondition.Enabled](Outlook.FormNameRuleCondition.Enabled.md)** would be **True**.

For more information on specifying rule actions, see [Specify Rule Conditions](../outlook/How-to/Rules/specifying-rule-conditions.md).


## Properties



|Name|
|:-----|
|[Application](Outlook.FormNameRuleCondition.Application.md)|
|[Class](Outlook.FormNameRuleCondition.Class.md)|
|[ConditionType](Outlook.FormNameRuleCondition.ConditionType.md)|
|[Enabled](Outlook.FormNameRuleCondition.Enabled.md)|
|[FormName](Outlook.FormNameRuleCondition.FormName.md)|
|[Parent](Outlook.FormNameRuleCondition.Parent.md)|
|[Session](Outlook.FormNameRuleCondition.Session.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]