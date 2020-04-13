---
title: RuleActions object (Outlook)
keywords: vbaol11.chm3162
f1_keywords:
- vbaol11.chm3162
ms.prod: outlook
api_name:
- Outlook.RuleActions
ms.assetid: 82ba76cd-86a4-3372-cb51-2df1d58c8b71
ms.date: 06/08/2017
localization_priority: Normal
---


# RuleActions object (Outlook)

The **RuleActions** object contains a set of **[RuleAction](Outlook.RuleAction.md)** objects or objects derived from **RuleAction**, representing the actions that are executed on a **[Rule](Outlook.Rule.md)** object.


## Remarks

The **RuleActions** object is a fixed collection. **RuleAction** objects or types that derive from the **RuleAction** object cannot be added to or removed from the **RuleActions** object.

The Rules object model provides partial parity with the Rules and Alerts Wizard in the Outlook user interface. It supports the most commonly used rule actions and conditions. Although it does not support creating rules with any rule action or rule condition that the Wizard supports, you can still enumerate and enable these rule actions and conditions in existing rules. 

For more information on rule actions, see [Specifying Rule Actions](../outlook/How-to/Rules/specifying-rule-actions.md) and [How to: Create a Rule to Move Specific Emails to a Folder](../outlook/How-to/Rules/create-a-rule-to-move-specific-e-mails-to-a-folder.md).


## Methods



|Name|
|:-----|
|[Item](Outlook.RuleActions.Item.md)|

## Properties



|Name|
|:-----|
|[Application](Outlook.RuleActions.Application.md)|
|[AssignToCategory](Outlook.RuleActions.AssignToCategory.md)|
|[CC](Outlook.RuleActions.CC.md)|
|[Class](Outlook.RuleActions.Class.md)|
|[ClearCategories](Outlook.RuleActions.ClearCategories.md)|
|[CopyToFolder](Outlook.RuleActions.CopyToFolder.md)|
|[Count](Outlook.RuleActions.Count.md)|
|[Delete](Outlook.RuleActions.Delete.md)|
|[DeletePermanently](Outlook.RuleActions.DeletePermanently.md)|
|[DesktopAlert](Outlook.RuleActions.DesktopAlert.md)|
|[Forward](Outlook.RuleActions.Forward.md)|
|[ForwardAsAttachment](Outlook.RuleActions.ForwardAsAttachment.md)|
|[MarkAsTask](Outlook.RuleActions.MarkAsTask.md)|
|[MoveToFolder](Outlook.RuleActions.MoveToFolder.md)|
|[NewItemAlert](Outlook.RuleActions.NewItemAlert.md)|
|[NotifyDelivery](Outlook.RuleActions.NotifyDelivery.md)|
|[NotifyRead](Outlook.RuleActions.NotifyRead.md)|
|[Parent](Outlook.RuleActions.Parent.md)|
|[PlaySound](Outlook.RuleActions.PlaySound.md)|
|[Redirect](Outlook.RuleActions.Redirect.md)|
|[Session](Outlook.RuleActions.Session.md)|
|[Stop](Outlook.RuleActions.Stop.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]