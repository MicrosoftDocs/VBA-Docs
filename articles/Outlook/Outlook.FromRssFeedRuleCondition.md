---
title: FromRssFeedRuleCondition Object (Outlook)
keywords: vbaol11.chm3260
f1_keywords:
- vbaol11.chm3260
ms.prod: outlook
api_name:
- Outlook.FromRssFeedRuleCondition
ms.assetid: 8de6e629-7e3d-b4df-d758-a5bff3abd6a1
ms.date: 06/08/2017
---


# FromRssFeedRuleCondition Object (Outlook)

Represents a rule condition that evaluates whether an item is from a specified RSS subscription.


## Remarks

 **FromRssFeedRuleCondition** is derived from the **[RuleCondition](Outlook.RuleCondition.md)** object. Each rule is associated with a **[RuleConditions](Outlook.RuleConditions.md)** object, which has a **[RuleConditions.FromRssFeed](Outlook.RuleConditions.FromRssFeed.md)** property. The **FromRssFeed** property always returns a **FromRssFeedRuleCondition** object. If the rule has any of these rule conditions enabled, then **[FromRssFeedRuleCondition.Enabled](Outlook.FromRssFeedRuleCondition.Enabled.md)** is **True**.

For more information about specifying rule actions, see [Specify Rule Conditions](http://msdn.microsoft.com/library/812c131a-fe23-1b8b-5e2d-9459d7102630%28Office.15%29.aspx).


## Properties



|**Name**|
|:-----|
|[Application](Outlook.FromRssFeedRuleCondition.Application.md)|
|[Class](Outlook.FromRssFeedRuleCondition.Class.md)|
|[ConditionType](Outlook.FromRssFeedRuleCondition.ConditionType.md)|
|[Enabled](Outlook.FromRssFeedRuleCondition.Enabled.md)|
|[FromRssFeed](Outlook.FromRssFeedRuleCondition.FromRssFeed.md)|
|[Parent](Outlook.FromRssFeedRuleCondition.Parent.md)|
|[Session](fromrssfeedrulecondition-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
