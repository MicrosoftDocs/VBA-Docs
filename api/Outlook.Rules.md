---
title: Rules object (Outlook)
keywords: vbaol11.chm3160
f1_keywords:
- vbaol11.chm3160
ms.prod: outlook
api_name:
- Outlook.Rules
ms.assetid: dd41b4de-bf5f-5532-46c9-394a5d078bec
ms.date: 06/08/2017
localization_priority: Normal
---


# Rules object (Outlook)

Represents a set of  **[Rule](Outlook.Rule.md)** objects that are the rules available in the current session.


## Remarks

The Rules object model consists primarily of these objects:  **Rules**, **Rule**, **[RuleActions](Outlook.RuleActions.md)**, **[RuleConditions](Outlook.RuleConditions.md)**, **[RuleAction](Outlook.RuleAction.md)**, **[RuleCondition](Outlook.RuleCondition.md)**, and the derived objects for certain rule actions and rule conditions. It provides partial parity with the Rules and Alerts Wizard in the Outlook user interface. Although it does not support creation of every single rule that you can possibly create using the Wizard, it supports the most commonly used rule actions and conditions.

For more information on how to programmatically create, edit, and delete rules, see [Managing Rules in the Outlook Object Model](../outlook/How-to/Rules/managing-rules-in-the-outlook-object-model.md) and [How to: Create a Rule to Move Specific Emails to a Folder](../outlook/How-to/Rules/create-a-rule-to-move-specific-e-mails-to-a-folder.md).


## Methods



|Name|
|:-----|
|[Create](Outlook.Rules.Create.md)|
|[Item](Outlook.Rules.Item.md)|
|[Remove](Outlook.Rules.Remove.md)|
|[Save](Outlook.Rules.Save.md)|

## Properties



|Name|
|:-----|
|[Application](Outlook.Rules.Application.md)|
|[Class](Outlook.Rules.Class.md)|
|[Count](Outlook.Rules.Count.md)|
|[IsRssRulesProcessingEnabled](Outlook.Rules.IsRssRulesProcessingEnabled.md)|
|[Parent](Outlook.Rules.Parent.md)|
|[Session](Outlook.Rules.Session.md)|

## See also


[Rules Object Members](overview/Outlook.md)
[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]