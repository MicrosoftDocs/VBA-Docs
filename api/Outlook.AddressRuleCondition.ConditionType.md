---
title: AddressRuleCondition.ConditionType property (Outlook)
keywords: vbaol11.chm2954
f1_keywords:
- vbaol11.chm2954
ms.prod: outlook
api_name:
- Outlook.AddressRuleCondition.ConditionType
ms.assetid: 8b531745-1a4d-d903-5c7d-465b9fd8cbf3
ms.date: 06/08/2017
localization_priority: Normal
---


# AddressRuleCondition.ConditionType property (Outlook)

Returns a constant from the  **[OlRuleConditionType](Outlook.OlRuleConditionType.md)** enumeration that indicates the type of rule condition. Read-only.


## Syntax

_expression_. `ConditionType`

_expression_ A variable that represents an [AddressRuleCondition](Outlook.AddressRuleCondition.md) object.


## Remarks

The  **[AddressRuleCondition](Outlook.AddressRuleCondition.md)** object is used by rules of types **olRuleSend** and **olRuleReceive**. If the rule is created as an **olRuleSend** rule, then the type of the associated **AddressRuleCondition** object will be **olConditionSenderAddress**. If the rule is created as an **olRuleReceive** rule, then the type of the associated **AddressRuleCondition** object will be **olConditionRecipientAddress**.

This however does not mean that the rule always has a defined rule condition for sender or recipient addresses. Regardless of whether there exists such a defined or enabled rule condition, the  **AddressRuleCondition.ConditionType** property is always initialized once the associated rule is created. For more information on rule conditions, see [Specify Rule Conditions](../outlook/How-to/Rules/specifying-rule-conditions.md).


## See also


[AddressRuleCondition Object](Outlook.AddressRuleCondition.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]