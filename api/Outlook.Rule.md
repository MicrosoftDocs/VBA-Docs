---
title: Rule object (Outlook)
keywords: vbaol11.chm3161
f1_keywords:
- vbaol11.chm3161
ms.prod: outlook
api_name:
- Outlook.Rule
ms.assetid: ea2ddbcc-fd65-a636-c6da-79950033f385
ms.date: 06/08/2017
localization_priority: Normal
---


# Rule object (Outlook)

Represents an Outlook rule.


## Remarks

Both client and server side rules are represented by the  **Rule** object.

The Rules object model consists primarily of these objects:  **[Rules](Outlook.Rules.md)**, **Rule**, **[RuleActions](Outlook.RuleActions.md)**, **[RuleConditions](Outlook.RuleConditions.md)**, **[RuleAction](Outlook.RuleAction.md)**, **[RuleCondition](Outlook.RuleCondition.md)**, and the derived objects for certain rule actions and rule conditions. It provides partial parity with the Rules and Alerts Wizard in the Outlook user interface. Although it does not support creation of every single rule that you can possibly create using the Wizard, it supports the most commonly used rule actions and conditions.

For more information on how to programmatically create, edit, and delete rules, see [Manage Rules in the Outlook Object Model](../outlook/How-to/Rules/managing-rules-in-the-outlook-object-model.md) and [How to: Create a Rule to Move Specific Emails to a Folder](../outlook/How-to/Rules/create-a-rule-to-move-specific-e-mails-to-a-folder.md).


## Methods



|Name|
|:-----|
|[Execute](Outlook.Rule.Execute.md)|

## Properties



|Name|
|:-----|
|[Actions](Outlook.Rule.Actions.md)|
|[Application](Outlook.Rule.Application.md)|
|[Class](Outlook.Rule.Class.md)|
|[Conditions](Outlook.Rule.Conditions.md)|
|[Enabled](Outlook.Rule.Enabled.md)|
|[Exceptions](Outlook.Rule.Exceptions.md)|
|[ExecutionOrder](Outlook.Rule.ExecutionOrder.md)|
|[IsLocalRule](Outlook.Rule.IsLocalRule.md)|
|[Name](Outlook.Rule.Name.md)|
|[Parent](Outlook.Rule.Parent.md)|
|[RuleType](Outlook.Rule.RuleType.md)|
|[Session](Outlook.Rule.Session.md)|

## See also


[Rule Object Members](overview/Outlook.md)
[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]