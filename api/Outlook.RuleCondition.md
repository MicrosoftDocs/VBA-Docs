---
title: RuleCondition object (Outlook)
keywords: vbaol11.chm3173
f1_keywords:
- vbaol11.chm3173
ms.prod: outlook
api_name:
- Outlook.RuleCondition
ms.assetid: e03f91c2-2c08-b036-104a-d6246f28bc2d
ms.date: 06/08/2017
localization_priority: Normal
---


# RuleCondition object (Outlook)

The  **RuleCondition** object represents either a condition that must be met before a rule executes, or an exception condition that must not be met before a rule executes.


## Remarks

 **RuleCondition** is the base class for rule conditions that are supported in programmatic rule creation. The classes derived from **RuleCondition** include:


-  **[AccountRuleCondition](Outlook.AccountRuleCondition.md)**
    
-  **[AddressRuleCondition](Outlook.AddressRuleCondition.md)**
    
-  **[CategoryRuleCondition](Outlook.CategoryRuleCondition.md)**
    
-  **[FromRssFeedRuleCondition](Outlook.FromRssFeedRuleCondition.md)**
    
-  **[FormNameRuleCondition](Outlook.FormNameRuleCondition.md)**
    
-  **[ImportanceRuleCondition](Outlook.ImportanceRuleCondition.md)**
    
-  **[SenderInAddressListRuleCondition](Outlook.SenderInAddressListRuleCondition.md)**
    
-  **[TextRuleCondition](Outlook.TextRuleCondition.md)**
    
-  **[ToOrFromRuleCondition](Outlook.ToOrFromRuleCondition.md)**
    


The Rules object model provides partial parity with the Rules and Alerts Wizard in the Outlook user interface. It supports the most commonly used rule actions and conditions. Although it does not support creating rules with each rule action or rule condition that the Wizard supports, you can still enumerate and enable these rule actions and conditions in existing rules. 

For more information on rule conditions, see [Specifying Rule Conditions](../outlook/How-to/Rules/specifying-rule-conditions.md) and [How to: Create a Rule to Move Specific Emails to a Folder](../outlook/How-to/Rules/create-a-rule-to-move-specific-e-mails-to-a-folder.md).


## Properties



|Name|
|:-----|
|[Application](Outlook.RuleCondition.Application.md)|
|[Class](Outlook.RuleCondition.Class.md)|
|[ConditionType](Outlook.RuleCondition.ConditionType.md)|
|[Enabled](Outlook.RuleCondition.Enabled.md)|
|[Parent](Outlook.RuleCondition.Parent.md)|
|[Session](Outlook.RuleCondition.Session.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]