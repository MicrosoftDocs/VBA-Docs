---
title: AddressRuleCondition object (Outlook)
keywords: vbaol11.chm3203
f1_keywords:
- vbaol11.chm3203
ms.prod: outlook
api_name:
- Outlook.AddressRuleCondition
ms.assetid: 8cf897ad-a8f9-67ea-c0fa-d7f4bb917bd4
ms.date: 06/08/2017
localization_priority: Normal
---


# AddressRuleCondition object (Outlook)

Represents a rule condition that evaluates whether the address for the recipient or sender of the message is contained in the address specified in  **[AddressRuleCondition.Address](Outlook.AddressRuleCondition.Address.md)**.


## Remarks

 **AddressRuleCondition** is derived from the **[RuleCondition](Outlook.RuleCondition.md)** object. Each rule is associated with a **[RuleConditions](Outlook.RuleConditions.md)** object which has a **[RecipientAddress](Outlook.RuleConditions.RecipientAddress.md)** property and a **[SenderAddress](Outlook.RuleConditions.SenderAddress.md)**. Each of these properties always returns a **AddressRuleCondition** object. **[AddressRuleCondition.ConditionType](Outlook.AddressRuleCondition.ConditionType.md)** distinguishes among these rule conditions. If the rule has any of these rule conditions enabled, then **[AddressRuleCondition.Enabled](Outlook.AddressRuleCondition.Enabled.md)** would be **True**.

For more information on specifying rule actions, see [Specifying Rule Conditions](../outlook/How-to/Rules/specifying-rule-conditions.md).


## Properties



|Name|
|:-----|
|[Address](Outlook.AddressRuleCondition.Address.md)|
|[Application](Outlook.AddressRuleCondition.Application.md)|
|[Class](Outlook.AddressRuleCondition.Class.md)|
|[ConditionType](Outlook.AddressRuleCondition.ConditionType.md)|
|[Enabled](Outlook.AddressRuleCondition.Enabled.md)|
|[Parent](Outlook.AddressRuleCondition.Parent.md)|
|[Session](Outlook.AddressRuleCondition.Session.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]