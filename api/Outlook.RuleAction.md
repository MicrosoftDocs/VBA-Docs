---
title: RuleAction object (Outlook)
keywords: vbaol11.chm3163
f1_keywords:
- vbaol11.chm3163
ms.prod: outlook
api_name:
- Outlook.RuleAction
ms.assetid: 6451788f-e5ed-239c-a34d-b564b52d8955
ms.date: 06/08/2017
localization_priority: Normal
---


# RuleAction object (Outlook)

Represents an action that is run when a  **[Rule](Outlook.Rule.md)** object executes.


## Remarks

 **RuleAction** is the base class for rule actions that are supported in programmatic rule creation. The classes derived from **RuleAction** include:


-  **[AssignToCategoryRuleAction](Outlook.AssignToCategoryRuleAction.md)**
    
-  **[MarkAsTaskRuleAction](Outlook.MarkAsTaskRuleAction.md)**
    
-  **[MoveOrCopyRuleAction](Outlook.MoveOrCopyRuleAction.md)**
    
-  **[NewItemAlertRuleAction](Outlook.NewItemAlertRuleAction.md)**
    
-  **[PlaySoundRuleAction](Outlook.PlaySoundRuleAction.md)**
    
-  **[SendRuleAction](Outlook.SendRuleAction.md)**
    


The Rules object model provides partial parity with the Rules and Alerts Wizard in the Outlook user interface. It supports the most commonly used rule actions and conditions. Although it does not support creating rules with each rule action or rule condition that the Wizard supports, you can still enumerate and enable these rule actions and conditions in existing rules. 

For more information on rule actions, see [Specifying Rule Actions](../outlook/How-to/Rules/specifying-rule-actions.md) and [How to: Create a Rule to Move Specific Emails to a Folder](../outlook/How-to/Rules/create-a-rule-to-move-specific-e-mails-to-a-folder.md).


## Properties



|Name|
|:-----|
|[ActionType](Outlook.RuleAction.ActionType.md)|
|[Application](Outlook.RuleAction.Application.md)|
|[Class](Outlook.RuleAction.Class.md)|
|[Enabled](Outlook.RuleAction.Enabled.md)|
|[Parent](Outlook.RuleAction.Parent.md)|
|[Session](Outlook.RuleAction.Session.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]