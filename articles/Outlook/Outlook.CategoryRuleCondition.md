---
title: CategoryRuleCondition Object (Outlook)
keywords: vbaol11.chm3179
f1_keywords:
- vbaol11.chm3179
ms.prod: outlook
api_name:
- Outlook.CategoryRuleCondition
ms.assetid: 7a9b8271-d673-1c69-9a2a-11fd1e5fb262
ms.date: 06/08/2017
---


# CategoryRuleCondition Object (Outlook)

Represents a rule condition that evaluates categories on a message as compared with  **CategoryRuleCondition.Categories**.


## Remarks

 **CategoryRuleCondition** is derived from the **[RuleCondition](Outlook.RuleCondition.md)** object. Each rule is associated with a **[RuleConditions](Outlook.RuleConditions.md)** object which has a **[RuleConditions.Category](Outlook.RuleConditions.Category.md)** property. The **Category** property always returns a **CategoryRuleCondition** object. If the rule has any of these rule conditions enabled, then **[CategoryRuleCondition.Enabled](Outlook.CategoryRuleCondition.Enabled.md)** would be **True**.

For more information on specifying rule actions, see [Specify Rule Conditions](http://msdn.microsoft.com/library/812c131a-fe23-1b8b-5e2d-9459d7102630%28Office.15%29.aspx).


## Properties



|**Name**|
|:-----|
|[Application](Outlook.CategoryRuleCondition.Application.md)|
|[Categories](Outlook.CategoryRuleCondition.Categories.md)|
|[Class](Outlook.CategoryRuleCondition.Class.md)|
|[ConditionType](Outlook.CategoryRuleCondition.ConditionType.md)|
|[Enabled](Outlook.CategoryRuleCondition.Enabled.md)|
|[Parent](Outlook.CategoryRuleCondition.Parent.md)|
|[Session](categoryrulecondition-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
