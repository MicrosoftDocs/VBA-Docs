---
title: AssignToCategoryRuleAction Object (Outlook)
keywords: vbaol11.chm3168
f1_keywords:
- vbaol11.chm3168
ms.prod: outlook
api_name:
- Outlook.AssignToCategoryRuleAction
ms.assetid: 402f4742-72ba-2559-4e4c-e2b8248cd7f6
ms.date: 06/08/2017
---


# AssignToCategoryRuleAction Object (Outlook)

Represents an action that assigns categories to a message.


## Remarks

 **AssignToCategoryRuleAction** is derived from the **[RuleAction](Outlook.RuleAction.md)** object. Each rule is associated with a **[RuleActions](Outlook.RuleActions.md)** object which has an **[AssignToCategory](Outlook.RuleActions.AssignToCategory.md)** property. The **AssignToCategory** property always returns an **[AssignToCategoryRuleAction](Outlook.AssignToCategoryRuleAction.md)** object. If the rule has an enabled rule action that assigns a message with some specified categories, then **[AssignToCategoryRuleAction.Enabled](Outlook.AssignToCategoryRuleAction.Enabled.md)** would be **True**.

For more information on specifying rule actions, see [Specify Rule Actions](http://msdn.microsoft.com/library/c5f83c81-0e01-38aa-5ec7-3932b4443e43%28Office.15%29.aspx).


## Properties



|**Name**|
|:-----|
|[ActionType](Outlook.AssignToCategoryRuleAction.ActionType.md)|
|[Application](Outlook.AssignToCategoryRuleAction.Application.md)|
|[Categories](Outlook.AssignToCategoryRuleAction.Categories.md)|
|[Class](Outlook.AssignToCategoryRuleAction.Class.md)|
|[Enabled](Outlook.AssignToCategoryRuleAction.Enabled.md)|
|[Parent](Outlook.AssignToCategoryRuleAction.Parent.md)|
|[Session](assigntocategoryruleaction-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
