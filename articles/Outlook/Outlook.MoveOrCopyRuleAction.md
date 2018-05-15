---
title: MoveOrCopyRuleAction Object (Outlook)
keywords: vbaol11.chm3164
f1_keywords:
- vbaol11.chm3164
ms.prod: outlook
api_name:
- Outlook.MoveOrCopyRuleAction
ms.assetid: db951ad8-0d05-1696-acf4-c1da4fbdee33
ms.date: 06/08/2017
---


# MoveOrCopyRuleAction Object (Outlook)

Represents an action that moves or copies a message.


## Remarks

 **MoveOrCopyRuleAction** is derived from the **[RuleAction](Outlook.RuleAction.md)** object. Each rule is associated with a **[RuleActions](Outlook.RuleActions.md)** object which has a **[CopyToFolder](Outlook.RuleActions.CopyToFolder.md)** property and a **[MoveToFolder](Outlook.RuleActions.MoveToFolder.md)** property. Each of these properties always returns a **MoveOrCopyRuleAction** object. **[MoveOrCopyRuleAction.ActionType](Outlook.MoveOrCopyRuleAction.ActionType.md)** distinguishes between the two action types. If the rule has an enabled rule action that copies or moves a message to the specified folder, then the corresponding **[MoveOrCopyRuleAction.Enabled](Outlook.MoveOrCopyRuleAction.Enabled.md)** would be **True**.

For more information on specifying rule actions, see [Specify Rule Actions](http://msdn.microsoft.com/library/c5f83c81-0e01-38aa-5ec7-3932b4443e43%28Office.15%29.aspx).


## Properties



|**Name**|
|:-----|
|[ActionType](Outlook.MoveOrCopyRuleAction.ActionType.md)|
|[Application](Outlook.MoveOrCopyRuleAction.Application.md)|
|[Class](Outlook.MoveOrCopyRuleAction.Class.md)|
|[Enabled](Outlook.MoveOrCopyRuleAction.Enabled.md)|
|[Folder](Outlook.MoveOrCopyRuleAction.Folder.md)|
|[Parent](moveorcopyruleaction-parent-property-outlook.md)|
|[Session](moveorcopyruleaction-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
