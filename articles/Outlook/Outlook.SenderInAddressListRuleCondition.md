---
title: SenderInAddressListRuleCondition Object (Outlook)
keywords: vbaol11.chm3182
f1_keywords:
- vbaol11.chm3182
ms.prod: outlook
api_name:
- Outlook.SenderInAddressListRuleCondition
ms.assetid: c43aa055-8d4f-e264-07dd-4c5519faf1c7
ms.date: 06/08/2017
---


# SenderInAddressListRuleCondition Object (Outlook)

Represents a rule condition that the sender's address is in the address list specified in  **[AddressRuleCondition.Address](Outlook.AddressRuleCondition.Address.md)**.


## Remarks

 **SenderInAddressListRuleCondition** is derived from the **[RuleCondition](Outlook.RuleCondition.md)** object. Each rule is associated with a **[RuleConditions](Outlook.RuleConditions.md)** object which has a **[SenderInAddressList](Outlook.RuleConditions.SenderInAddressList.md)** property. The **SenderInAddressList** property always returns a **SenderInAddressListRuleCondition** object. If the rule has any of these rule conditions enabled, then **[SenderInAddressListRuleCondition.Enabled](Outlook.SenderInAddressListRuleCondition.Enabled.md)** would be **True**.

For more information on specifying rule conditions, see [Specify Rule Conditions](http://msdn.microsoft.com/library/812c131a-fe23-1b8b-5e2d-9459d7102630%28Office.15%29.aspx).


## Properties



|**Name**|
|:-----|
|[AddressList](Outlook.SenderInAddressListRuleCondition.AddressList.md)|
|[Application](Outlook.SenderInAddressListRuleCondition.Application.md)|
|[Class](Outlook.SenderInAddressListRuleCondition.Class.md)|
|[ConditionType](Outlook.SenderInAddressListRuleCondition.ConditionType.md)|
|[Enabled](Outlook.SenderInAddressListRuleCondition.Enabled.md)|
|[Parent](Outlook.SenderInAddressListRuleCondition.Parent.md)|
|[Session](senderinaddresslistrulecondition-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
