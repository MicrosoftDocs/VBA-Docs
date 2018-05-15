---
title: SendRuleAction Object (Outlook)
keywords: vbaol11.chm3165
f1_keywords:
- vbaol11.chm3165
ms.prod: outlook
api_name:
- Outlook.SendRuleAction
ms.assetid: 4ea8f519-8bb3-b0bf-9742-8a492e7ffff7
ms.date: 06/08/2017
---


# SendRuleAction Object (Outlook)

Represents an action that sends a message to one or more recipients.


## Remarks

 **SendRuleAction** is derived from the **[RuleAction](Outlook.RuleAction.md)** object. Each rule is associated with a **[RuleActions](Outlook.RuleActions.md)** object which has a **[CC](Outlook.RuleActions.CC.md)** property, a **[Forward](Outlook.RuleActions.Forward.md)** property, a **[ForwardAsAttachment](Outlook.RuleActions.ForwardAsAttachment.md)** property, and a **[Redirect](Outlook.RuleActions.Redirect.md)** property. Each of these properties always returns a **SendRuleAction** object. **[SendRuleAction.ActionType](Outlook.SendRuleAction.ActionType.md)** distinguishes among these rule actions. If the rule has any of the above rule actions enabled, then the **[Enabled](Outlook.SendRuleAction.Enabled.md)** property of the corresponding **SendRuleAction** object would be **True**.

For more information on specifying rule actions, see [Specify Rule Actions](http://msdn.microsoft.com/library/c5f83c81-0e01-38aa-5ec7-3932b4443e43%28Office.15%29.aspx).


## Properties



|**Name**|
|:-----|
|[ActionType](Outlook.SendRuleAction.ActionType.md)|
|[Application](Outlook.SendRuleAction.Application.md)|
|[Class](Outlook.SendRuleAction.Class.md)|
|[Enabled](Outlook.SendRuleAction.Enabled.md)|
|[Parent](Outlook.SendRuleAction.Parent.md)|
|[Recipients](sendruleaction-recipients-property-outlook.md)|
|[Session](sendruleaction-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
