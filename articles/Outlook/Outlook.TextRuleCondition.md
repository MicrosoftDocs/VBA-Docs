---
title: TextRuleCondition Object (Outlook)
keywords: vbaol11.chm3183
f1_keywords:
- vbaol11.chm3183
ms.prod: outlook
api_name:
- Outlook.TextRuleCondition
ms.assetid: 87e9ca00-7577-02c2-fb6f-a5dc2054ad8b
ms.date: 06/08/2017
---


# TextRuleCondition Object (Outlook)

Represents a rule condition that the part of the message, which can be the body, header, or subject, as specified by  **[TextRuleCondition.ConditionType](Outlook.TextRuleCondition.ConditionType.md)**, contains the words specified in **[TextRuleCondition.Text](Outlook.TextRuleCondition.Text.md)**.


## Remarks

 **TextRuleCondition** is derived from the **[RuleCondition](Outlook.RuleCondition.md)** object. Each rule is associated with a **[RuleConditions](Outlook.RuleConditions.md)** object which has the following properties: **[Body](Outlook.RuleConditions.Body.md)**, **[BodyOrSubject](Outlook.RuleConditions.BodyOrSubject.md)**, **[MessageHeader](Outlook.RuleConditions.MessageHeader.md)**, and **[Subject](Outlook.RuleConditions.Subject.md)**. Each of these properties always returns a **TextRuleCondition** object. **TextRuleCondition.ConditionType** distinguishes among these rule conditions. If the rule has any of these rule conditions enabled, then **[TextRuleCondition.Enabled](Outlook.TextRuleCondition.Enabled.md)** would be **True**.

For more information on specifying rule conditions, see [Specify Rule Conditions](http://msdn.microsoft.com/library/812c131a-fe23-1b8b-5e2d-9459d7102630%28Office.15%29.aspx).


## Properties



|**Name**|
|:-----|
|[Application](Outlook.TextRuleCondition.Application.md)|
|[Class](Outlook.TextRuleCondition.Class.md)|
|[ConditionType](Outlook.TextRuleCondition.ConditionType.md)|
|[Enabled](Outlook.TextRuleCondition.Enabled.md)|
|[Parent](textrulecondition-parent-property-outlook.md)|
|[Session](textrulecondition-session-property-outlook.md)|
|[Text](Outlook.TextRuleCondition.Text.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
