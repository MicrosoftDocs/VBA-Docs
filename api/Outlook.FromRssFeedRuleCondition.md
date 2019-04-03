---
title: FromRssFeedRuleCondition object (Outlook)
keywords: vbaol11.chm3260
f1_keywords:
- vbaol11.chm3260
ms.prod: outlook
api_name:
- Outlook.FromRssFeedRuleCondition
ms.assetid: 8de6e629-7e3d-b4df-d758-a5bff3abd6a1
ms.date: 06/08/2017
localization_priority: Normal
---


# FromRssFeedRuleCondition object (Outlook)

Represents a rule condition that evaluates whether an item is from a specified RSS subscription.


## Remarks

 **FromRssFeedRuleCondition** is derived from the **[RuleCondition](Outlook.RuleCondition.md)** object. Each rule is associated with a **[RuleConditions](Outlook.RuleConditions.md)** object, which has a **[RuleConditions.FromRssFeed](Outlook.RuleConditions.FromRssFeed.md)** property. The **FromRssFeed** property always returns a **FromRssFeedRuleCondition** object. If the rule has any of these rule conditions enabled, then **[FromRssFeedRuleCondition.Enabled](Outlook.FromRssFeedRuleCondition.Enabled.md)** is **True**.

For more information about specifying rule actions, see [Specify Rule Conditions](../outlook/How-to/Rules/specifying-rule-conditions.md).


## Properties



|Name|
|:-----|
|[Application](Outlook.FromRssFeedRuleCondition.Application.md)|
|[Class](Outlook.FromRssFeedRuleCondition.Class.md)|
|[ConditionType](Outlook.FromRssFeedRuleCondition.ConditionType.md)|
|[Enabled](Outlook.FromRssFeedRuleCondition.Enabled.md)|
|[FromRssFeed](Outlook.FromRssFeedRuleCondition.FromRssFeed.md)|
|[Parent](Outlook.FromRssFeedRuleCondition.Parent.md)|
|[Session](Outlook.FromRssFeedRuleCondition.Session.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]