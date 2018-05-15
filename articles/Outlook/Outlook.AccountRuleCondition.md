---
title: AccountRuleCondition Object (Outlook)
keywords: vbaol11.chm3175
f1_keywords:
- vbaol11.chm3175
ms.prod: outlook
api_name:
- Outlook.AccountRuleCondition
ms.assetid: 1b746449-1357-36c2-5081-392ea85fb71e
ms.date: 06/08/2017
---


# AccountRuleCondition Object (Outlook)

Represents a rule condition that evaluates whether an account was used to send a message.


## Remarks

 **AccountRuleCondition** is derived from the **[RuleCondition](Outlook.RuleCondition.md)** object. Each rule is associated with a **[RuleConditions](Outlook.RuleConditions.md)** object which has an **[Account](Outlook.RuleConditions.Account.md)** property. The **Account** property always returns a **AccountRuleCondition** object. If the rule has an enabled rule condition that the message is sent using a specified account, then **[AccountRuleCondition.Enabled](Outlook.AccountRuleCondition.Enabled.md)** would be **True**.

For more information on specifying rule actions, see [Specify Rule Conditions](http://msdn.microsoft.com/library/812c131a-fe23-1b8b-5e2d-9459d7102630%28Office.15%29.aspx).


## Properties



|**Name**|
|:-----|
|[Account](Outlook.AccountRuleCondition.Account.md)|
|[Application](Outlook.AccountRuleCondition.Application.md)|
|[Class](Outlook.AccountRuleCondition.Class.md)|
|[ConditionType](Outlook.AccountRuleCondition.ConditionType.md)|
|[Enabled](Outlook.AccountRuleCondition.Enabled.md)|
|[Parent](Outlook.AccountRuleCondition.Parent.md)|
|[Session](accountrulecondition-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
