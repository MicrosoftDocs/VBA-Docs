---
title: MarkAsTaskRuleAction Object (Outlook)
keywords: vbaol11.chm3170
f1_keywords:
- vbaol11.chm3170
ms.prod: outlook
api_name:
- Outlook.MarkAsTaskRuleAction
ms.assetid: 639d9242-7387-2b25-9d0f-f7a14cf16790
ms.date: 06/08/2017
---


# MarkAsTaskRuleAction Object (Outlook)

Represents an action that marks a message as a task.


## Remarks

 **MarkAsTaskRuleAction** is derived from the **[RuleAction](Outlook.RuleAction.md)** object. Each rule is associated with a **[RuleActions](Outlook.RuleActions.md)** object which has a **[MarkAsTask](Outlook.RuleActions.MarkAsTask.md)** property. The **MarkAsTask** property always returns a **MarkAsTaskRuleAction** object. If the rule has an enabled rule action that marks a message as a task, then **[MarkAsTaskRuleAction.Enabled](Outlook.MarkAsTaskRuleAction.Enabled.md)** would be **True**.

For more information on specifying rule actions, see [Specify Rule Actions](http://msdn.microsoft.com/library/c5f83c81-0e01-38aa-5ec7-3932b4443e43%28Office.15%29.aspx).


## Properties



|**Name**|
|:-----|
|[ActionType](Outlook.MarkAsTaskRuleAction.ActionType.md)|
|[Application](Outlook.MarkAsTaskRuleAction.Application.md)|
|[Class](Outlook.MarkAsTaskRuleAction.Class.md)|
|[Enabled](Outlook.MarkAsTaskRuleAction.Enabled.md)|
|[FlagTo](Outlook.MarkAsTaskRuleAction.FlagTo.md)|
|[MarkInterval](Outlook.MarkAsTaskRuleAction.MarkInterval.md)|
|[Parent](Outlook.MarkAsTaskRuleAction.Parent.md)|
|[Session](markastaskruleaction-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
