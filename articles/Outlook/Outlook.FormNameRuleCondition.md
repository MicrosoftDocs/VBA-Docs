---
title: FormNameRuleCondition Object (Outlook)
keywords: vbaol11.chm3180
f1_keywords:
- vbaol11.chm3180
ms.prod: outlook
api_name:
- Outlook.FormNameRuleCondition
ms.assetid: 75b7f687-66e6-4863-b8aa-f19e98fedc45
ms.date: 06/08/2017
---


# FormNameRuleCondition Object (Outlook)

Represents a rule condition that evaluates whether a form name was used to send or receive an item.


## Remarks

 **FormNameRuleCondition** is derived from the **[RuleCondition](Outlook.RuleCondition.md)** object. Each rule is associated with a **[RuleConditions](Outlook.RuleConditions.md)** object which has a **[FormName](Outlook.RuleConditions.FormName.md)** property. The **FormName** property always returns a **FormNameRuleCondition** object. If the rule has any of these rule conditions enabled, then **[FormNameRuleCondition.Enabled](Outlook.FormNameRuleCondition.Enabled.md)** would be **True**.

For more information on specifying rule actions, see [Specify Rule Conditions](http://msdn.microsoft.com/library/812c131a-fe23-1b8b-5e2d-9459d7102630%28Office.15%29.aspx).


## Properties



|**Name**|
|:-----|
|[Application](Outlook.FormNameRuleCondition.Application.md)|
|[Class](Outlook.FormNameRuleCondition.Class.md)|
|[ConditionType](Outlook.FormNameRuleCondition.ConditionType.md)|
|[Enabled](Outlook.FormNameRuleCondition.Enabled.md)|
|[FormName](Outlook.FormNameRuleCondition.FormName.md)|
|[Parent](Outlook.FormNameRuleCondition.Parent.md)|
|[Session](formnamerulecondition-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
