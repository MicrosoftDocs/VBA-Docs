---
title: TextRuleCondition object (Outlook)
keywords: vbaol11.chm3183
f1_keywords:
- vbaol11.chm3183
ms.prod: outlook
api_name:
- Outlook.TextRuleCondition
ms.assetid: 87e9ca00-7577-02c2-fb6f-a5dc2054ad8b
ms.date: 06/08/2017
localization_priority: Normal
---


# TextRuleCondition object (Outlook)

Represents a rule condition that the part of the message, which can be the body, header, or subject, as specified by  **[TextRuleCondition.ConditionType](Outlook.TextRuleCondition.ConditionType.md)**, contains the words specified in **[TextRuleCondition.Text](Outlook.TextRuleCondition.Text.md)**.


## Remarks

 **TextRuleCondition** is derived from the **[RuleCondition](Outlook.RuleCondition.md)** object. Each rule is associated with a **[RuleConditions](Outlook.RuleConditions.md)** object which has the following properties: **[Body](Outlook.RuleConditions.Body.md)**, **[BodyOrSubject](Outlook.RuleConditions.BodyOrSubject.md)**, **[MessageHeader](Outlook.RuleConditions.MessageHeader.md)**, and **[Subject](Outlook.RuleConditions.Subject.md)**. Each of these properties always returns a **TextRuleCondition** object. **TextRuleCondition.ConditionType** distinguishes among these rule conditions. If the rule has any of these rule conditions enabled, then **[TextRuleCondition.Enabled](Outlook.TextRuleCondition.Enabled.md)** would be **True**.

For more information on specifying rule conditions, see [Specify Rule Conditions](../outlook/How-to/Rules/specifying-rule-conditions.md).


## Properties



|Name|
|:-----|
|[Application](Outlook.TextRuleCondition.Application.md)|
|[Class](Outlook.TextRuleCondition.Class.md)|
|[ConditionType](Outlook.TextRuleCondition.ConditionType.md)|
|[Enabled](Outlook.TextRuleCondition.Enabled.md)|
|[Parent](Outlook.TextRuleCondition.Parent.md)|
|[Session](Outlook.TextRuleCondition.Session.md)|
|[Text](Outlook.TextRuleCondition.Text.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]