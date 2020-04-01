---
title: Specifying Rule Actions
ms.prod: outlook
ms.assetid: c5f83c81-0e01-38aa-5ec7-3932b4443e43
ms.date: 06/08/2017
localization_priority: Normal
---


# Specifying Rule Actions

The Rules object model supports the most commonly used rule actions and conditions. Each **Rule](../../../api/Outlook.Rule.md)** object has an **[Actions](../../../api/Outlook.Rule.Actions.md)** property that represents the rule actions for that rule, as well as a **[Conditions](../../../api/Outlook.Rule.Conditions.md)** property and an **[Exceptions](../../../api/Outlook.Rule.Exceptions.md)** property that represent the conditions for that rule. This topic describes how the Rules object model supports rule actions.

Rule actions for a rule are represented by a **RuleActions](../../../api/Outlook.RuleActions.md)** collection object. A **RuleActions** object has properties that correspond to each commonly used rule action in a rule. For example, if a rule specifies two actions - moving the message to a specific folder and plays a sound - then the **[MoveToFolder](../../../api/Outlook.RuleActions.MoveToFolder.md)** and **[PlaySound](../../../api/Outlook.RuleActions.PlaySound.md)** properties of the rule's **RuleActions** collection object will return respective rule action objects that are enabled (**[RuleAction.Enabled](../../../api/Outlook.RuleAction.Enabled.md)** is **True**). 

Actions that are not specified in a rule will not be enabled in the corresponding **uleAction** object (**RuleAction.Enabled** is **False**). These rule action objects are represented by either the  **leAction** object or customized objects derived from the **RuleAction** object. In the last example, specifically, the **RuleActions.MoveToFolder** property will return a **[MoveOrCopyRuleAction](../../../api/Outlook.MoveOrCopyRuleAction.md)** object, and the **RuleActions.PlaySound** property will return a **[PlaySoundRuleAction](../../../api/Outlook.PlaySoundRuleAction.md)** object, both of which are derived from the **RuleAction** object. The **RuleAction** object and its derived objects have the **ActionType** property that will indicate the type of the rule action. For example, **[MoveOrCopyRuleAction.ActionType](../../../api/Outlook.MoveOrCopyRuleAction.ActionType.md)** will indicate the value **olRuleActionMoveToFolder**, and * **aySoundRuleAction.ActionType](../../../api/Outlook.PlaySoundRuleAction.ActionType.md)** will indicate **olRuleActionPlay**. 

Note that the Rules object model maintains partial parity with the Rules and Alerts Wizard. This means that while you can use the Wizard to create rules that specify any action and condition that you see in the Wizard, you can programmatically create rules that use some but not all of these actions and conditions. An example of an action that the Rules object model supports for rules created by the Wizard but not for those created by the object model is requesting a server reply. You can use the Wizard to create a rule specifying a certain server reply as an action. 

Using the Rules object model, you can enumerate these kinds of rules in the **ules** collection - for each rule in the **Rules** collection, enumerate its **RuleActions** collection and look for an enabled rule action for a server reply. In code, this would mean for each rule in the **Rules** collection, enumerate **[RuleActions.Item(Index)](../../../api/Outlook.RuleActions.Item.md)** using the _Index_ from 1 to **[RuleActions.Count](../../../api/Outlook.RuleActions.Count.md)**, and look for an enabled action with  **tionType** equal to **olRuleActionServerReply**. You can also enable or disable such a rule action in a rule. However, you cannot programmatically create a rule that specifies the * **uleActionServerReply** action.

The following table lists all the rule actions supported by the Rules and Alerts Wizard, and whether each rule action is supported when creating a rule using the Rules object model. A rule action that is not supported in rules created by the Rules object model is supported only for programmatic enumeration and enabling or disabling in existing rules created by the Rules and Alerts Wizard. The table also shows whether the rule action applies to rules with the **lRuleReceive** rule type or **olRuleSend** rule type, or both.


| **Action**| **Constant in olRuleActionType**| **Supported when creating new rules programmatically?**| **Apply to olRuleReceive rules?**| **Apply to olRuleSend rules?**|
|:-----|:-----|:-----|:-----|:-----|
|Assign the message to the categories specified in the **AssignToCategoryRuleAction.Categories](../../../api/Outlook.AssignToCategoryRuleAction.Categories.md)** property| **olRuleActionAssignToCategory**|Yes|Yes|Yes|
|Cc the message to the recipient list specified in the **SendRuleAction.Recipients](../../../api/Outlook.SendRuleAction.Recipients.md)** property| **olRuleActionCcMessage**|Yes|No|Yes|
|Clear all categories for the message.| **olRuleActionClearCategories**|Yes|Yes|Yes|
|Copy the message to folder specified in the **[MoveOrCopyRuleAction.Folder](../../../api/Outlook.MoveOrCopyRuleAction.Folder.md)** property| **olRuleActionCopyToFolder**|Yes|Yes|Yes|
|Run a custom action| **olRuleActionCustomAction**|No|Yes|Yes|
|Defer the delivery by a specified number of minutes| **olRuleActionDefer**|No|No|Yes|
|Delete the message| **olRuleActionDelete**|Yes|Yes|No|
|Permanently delete the message| **olRuleActionDeletePermanently**|Yes|Yes|No|
|Display a desktop alert| **olRuleActionDesktopAlert**|Yes|Yes|No|
|Clear the message flag| **olRuleActionFlagClear**|No|Yes|No|
|Flag the message with the color specified | **olRuleActionFlagColor**|No|Yes|No|
|Flag the message for action in days specified | **olRuleActionFlagForActionInDays**|No|Yes|Yes|
|Forward the message to the recipient list specified in the **endRuleAction.Recipients** property| **olRuleActionForward**|Yes|Yes|No|
|Forward the message as an attachment to the recipient list specified in the **endRuleAction.Recipients** property| **olRuleActionForwardAsAttachment**|Yes|Yes|No|
|Mark the message with the specified Importance| **olRuleActionImportance**|No|Yes|Yes|
|Mark message as a task for followup using the **FlagTo](../../../api/Outlook.MarkAsTaskRuleAction.FlagTo.md)** and **[MarkInterval](../../../api/Outlook.MarkAsTaskRuleAction.MarkInterval.md)** properties of the **[MarkAsTaskRuleAction](../../../api/Outlook.MarkAsTaskRuleAction.md)** object| **olRuleActionMarkAsTask**|Yes|Yes|No|
|Mark as read| **olRuleActionMarkRead**|No|Yes|No|
|Move the message to the folder specified in the **oveOrCopyRuleAction.Folder** property| **olRuleActionMoveToFolder**|Yes|Yes|No|
|Display the message specified in the **NewItemAlertRuleAction.Text](../../../api/Outlook.NewItemAlertRuleAction.Text.md)** property| **olRuleActionNewItemAlert**|Yes|Yes|No|
|Notify that the message has been delivered| **olRuleActionNotifyDelivery**|Yes|No|Yes|
|Notify that the message has been read| **olRuleActionNotifyRead**|Yes|No|Yes|
|Play the .wav file specified in the **PlaySoundRuleAction.FilePath](../../../api/Outlook.PlaySoundRuleAction.FilePath.md)** property| **olRuleActionPlaysound**|Yes|Yes|No|
|Print the message to the default printer| **olRuleActionPrint**|No|Yes|No|
|Redirect the message to the recipient list specified in the **endRuleAction.Recipients** property| **olRuleActionRedirect**|Yes|Yes|No|
|Start a script| **olRuleActionRunScript**|No|Yes|No|
|Mark the message with the specified sensitivity| **olRuleActionSensitivity**|No|No|Yes|
|Have server reply using the specified message | **olRuleActionServerReply**|No|Yes|No|
|Start an .exe| **olRuleActionStartApplication**|No|Yes|No|
|Stop processing more rules| **olRuleActionStop**|Yes|Yes|Yes|
|Reply using the specified template (.oft) file| **olRuleActionTemplate**|No|Yes|No|
|Unrecognized rule action| **olRuleActionUnknown**|No|Yes|No|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]