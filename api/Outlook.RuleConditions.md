---
title: RuleConditions object (Outlook)
keywords: vbaol11.chm3172
f1_keywords:
- vbaol11.chm3172
ms.prod: outlook
api_name:
- Outlook.RuleConditions
ms.assetid: e8e9a05a-b36b-add2-b294-8cdc5a97e119
ms.date: 06/08/2017
localization_priority: Normal
---


# RuleConditions object (Outlook)

Contains a set of  **[RuleCondition](Outlook.RuleCondition.md)** objects or objects derived from **RuleCondition**, representing the conditions or exception conditions that must be satisfied in order for the **[Rule](Outlook.Rule.md)** to execute.


## Remarks

The **RuleConditions** object include both rule conditions and rule exceptions. The type of rule condition that can be added to a **RuleConditions** collection depends upon the **[Rule.RuleType](Outlook.Rule.RuleType.md)**.

The **RuleConditions** object is a fixed collection. A **RuleCondition** object or a type that is derived from the **RuleCondition** object cannot be added or removed from the **RuleConditions** object.

The Rules object model provides partial parity with the Rules and Alerts Wizard in the Outlook user interface. It supports the most commonly used rule actions and conditions. Although it does not support creating rules with any rule action or rule condition that the Wizard supports, you can still enumerate and enable these rule actions and conditions in existing rules. 

For more information on rule conditions, see [Specifying Rule Conditions](../outlook/How-to/Rules/specifying-rule-conditions.md) and [How to: Create a Rule to Move Specific Emails to a Folder](../outlook/How-to/Rules/create-a-rule-to-move-specific-e-mails-to-a-folder.md).


## Methods



|Name|
|:-----|
|[Item](Outlook.RuleConditions.Item.md)|

## Properties



|Name|
|:-----|
|[Account](Outlook.RuleConditions.Account.md)|
|[AnyCategory](Outlook.RuleConditions.AnyCategory.md)|
|[Application](Outlook.RuleConditions.Application.md)|
|[Body](Outlook.RuleConditions.Body.md)|
|[BodyOrSubject](Outlook.RuleConditions.BodyOrSubject.md)|
|[Category](Outlook.RuleConditions.Category.md)|
|[CC](Outlook.RuleConditions.CC.md)|
|[Class](Outlook.RuleConditions.Class.md)|
|[Count](Outlook.RuleConditions.Count.md)|
|[FormName](Outlook.RuleConditions.FormName.md)|
|[From](Outlook.RuleConditions.From.md)|
|[FromAnyRSSFeed](Outlook.RuleConditions.FromAnyRSSFeed.md)|
|[FromRssFeed](Outlook.RuleConditions.FromRssFeed.md)|
|[HasAttachment](Outlook.RuleConditions.HasAttachment.md)|
|[Importance](Outlook.RuleConditions.Importance.md)|
|[MeetingInviteOrUpdate](Outlook.RuleConditions.MeetingInviteOrUpdate.md)|
|[MessageHeader](Outlook.RuleConditions.MessageHeader.md)|
|[NotTo](Outlook.RuleConditions.NotTo.md)|
|[OnLocalMachine](Outlook.RuleConditions.OnLocalMachine.md)|
|[OnlyToMe](Outlook.RuleConditions.OnlyToMe.md)|
|[OnOtherMachine](Outlook.RuleConditions.OnOtherMachine.md)|
|[Parent](Outlook.RuleConditions.Parent.md)|
|[RecipientAddress](Outlook.RuleConditions.RecipientAddress.md)|
|[SenderAddress](Outlook.RuleConditions.SenderAddress.md)|
|[SenderInAddressList](Outlook.RuleConditions.SenderInAddressList.md)|
|[SentTo](Outlook.RuleConditions.SentTo.md)|
|[Session](Outlook.RuleConditions.Session.md)|
|[Subject](Outlook.RuleConditions.Subject.md)|
|[ToMe](Outlook.RuleConditions.ToMe.md)|
|[ToOrCc](Outlook.RuleConditions.ToOrCc.md)|

## See also


[RuleConditions Object Members](overview/Outlook.md)
[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]