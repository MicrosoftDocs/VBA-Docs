---
title: ToOrFromRuleCondition Object (Outlook)
keywords: vbaol11.chm3181
f1_keywords:
- vbaol11.chm3181
ms.prod: outlook
api_name:
- Outlook.ToOrFromRuleCondition
ms.assetid: ec5cae2a-cde8-5681-6a49-74e2f0226a4f
ms.date: 06/08/2017
---


# ToOrFromRuleCondition Object (Outlook)

Represents a rule condition that the sender or the recipeints of the message, as specified by  **[ToOrFromRuleCondition.ConditionType](Outlook.ToOrFromRuleCondition.ConditionType.md)**, is in the recipients list specified in **[ToOrFromRuleCondition.Recipients](Outlook.ToOrFromRuleCondition.Recipients.md)**.


## Remarks

 **ToOrFromRuleCondition** is derived from the **[RuleCondition](Outlook.RuleCondition.md)** object. Each rule is associated with a **[RuleConditions](Outlook.RuleConditions.md)** object which has a **[SentTo](Outlook.RuleConditions.SentTo.md)** property and a **[From](Outlook.RuleConditions.From.md)**. Each of these properties always returns a **ToOrFromRuleCondition** object. **ToOrFromRuleCondition.ConditionType** distinguishes between these rule conditions. If the rule has any of these rule conditions enabled, then **[ToOrFromRuleCondition.Enabled](Outlook.ToOrFromRuleCondition.Enabled.md)** would be **True**.

For more information on specifying rule conditions, see [Specify Rule Conditions](http://msdn.microsoft.com/library/812c131a-fe23-1b8b-5e2d-9459d7102630%28Office.15%29.aspx).


## Properties



|**Name**|
|:-----|
|[Application](Outlook.ToOrFromRuleCondition.Application.md)|
|[Class](Outlook.ToOrFromRuleCondition.Class.md)|
|[ConditionType](Outlook.ToOrFromRuleCondition.ConditionType.md)|
|[Enabled](Outlook.ToOrFromRuleCondition.Enabled.md)|
|[Parent](toorfromrulecondition-parent-property-outlook.md)|
|[Recipients](Outlook.ToOrFromRuleCondition.Recipients.md)|
|[Session](toorfromrulecondition-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
