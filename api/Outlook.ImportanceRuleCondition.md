---
title: ImportanceRuleCondition object (Outlook)
keywords: vbaol11.chm3174
f1_keywords:
- vbaol11.chm3174
ms.prod: outlook
api_name:
- Outlook.ImportanceRuleCondition
ms.assetid: 52985055-f995-5613-d27f-7ad9618cfb46
ms.date: 06/08/2017
localization_priority: Normal
---


# ImportanceRuleCondition object (Outlook)

Represents a rule condition that evaluates the importance of a message.


## Remarks

 **ImportanceRuleCondition** is derived from the **[RuleCondition](Outlook.RuleCondition.md)** object. Each rule is associated with a **[RuleConditions](Outlook.RuleConditions.md)** object which has an **[Importance](Outlook.RuleConditions.Importance.md)** property. The **Importance** property always returns a **ImportanceRuleCondition** object. If the rule has any of these rule conditions enabled, then **[ImportanceRuleCondition.Enabled](Outlook.ImportanceRuleCondition.Enabled.md)** would be **True**.

For more information on specifying rule actions, see [Specify Rule Conditions](../outlook/How-to/Rules/specifying-rule-conditions.md).


## Properties



|Name|
|:-----|
|[Application](Outlook.ImportanceRuleCondition.Application.md)|
|[Class](Outlook.ImportanceRuleCondition.Class.md)|
|[ConditionType](Outlook.ImportanceRuleCondition.ConditionType.md)|
|[Enabled](Outlook.ImportanceRuleCondition.Enabled.md)|
|[Importance](Outlook.ImportanceRuleCondition.Importance.md)|
|[Parent](Outlook.ImportanceRuleCondition.Parent.md)|
|[Session](Outlook.ImportanceRuleCondition.Session.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]