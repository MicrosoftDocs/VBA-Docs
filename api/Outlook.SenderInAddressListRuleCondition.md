---
title: SenderInAddressListRuleCondition object (Outlook)
keywords: vbaol11.chm3182
f1_keywords:
- vbaol11.chm3182
ms.prod: outlook
api_name:
- Outlook.SenderInAddressListRuleCondition
ms.assetid: c43aa055-8d4f-e264-07dd-4c5519faf1c7
ms.date: 06/08/2017
localization_priority: Normal
---


# SenderInAddressListRuleCondition object (Outlook)

Represents a rule condition that the sender's address is in the address list specified in  **[AddressRuleCondition.Address](Outlook.AddressRuleCondition.Address.md)**.


## Remarks

 **SenderInAddressListRuleCondition** is derived from the **[RuleCondition](Outlook.RuleCondition.md)** object. Each rule is associated with a **[RuleConditions](Outlook.RuleConditions.md)** object which has a **[SenderInAddressList](Outlook.RuleConditions.SenderInAddressList.md)** property. The **SenderInAddressList** property always returns a **SenderInAddressListRuleCondition** object. If the rule has any of these rule conditions enabled, then **[SenderInAddressListRuleCondition.Enabled](Outlook.SenderInAddressListRuleCondition.Enabled.md)** would be **True**.

For more information on specifying rule conditions, see [Specify Rule Conditions](../outlook/How-to/Rules/specifying-rule-conditions.md).


## Properties



|Name|
|:-----|
|[AddressList](Outlook.SenderInAddressListRuleCondition.AddressList.md)|
|[Application](Outlook.SenderInAddressListRuleCondition.Application.md)|
|[Class](Outlook.SenderInAddressListRuleCondition.Class.md)|
|[ConditionType](Outlook.SenderInAddressListRuleCondition.ConditionType.md)|
|[Enabled](Outlook.SenderInAddressListRuleCondition.Enabled.md)|
|[Parent](Outlook.SenderInAddressListRuleCondition.Parent.md)|
|[Session](Outlook.SenderInAddressListRuleCondition.Session.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]