---
title: NewItemAlertRuleAction Object (Outlook)
keywords: vbaol11.chm3171
f1_keywords:
- vbaol11.chm3171
ms.prod: outlook
api_name:
- Outlook.NewItemAlertRuleAction
ms.assetid: 01d30816-50aa-ff23-69a0-4aa627b3d7e4
ms.date: 06/08/2017
---


# NewItemAlertRuleAction Object (Outlook)

Represents an action that displays a new item alert to the user.


## Remarks

 **NewItemAlertRuleAction** is derived from the **[RuleAction](Outlook.RuleAction.md)** object. Each rule is associated with a **[RuleActions](Outlook.RuleActions.md)** object which has a **[NewItemAlert](Outlook.RuleActions.NewItemAlert.md)** property. The **NewItemAlert** property always returns a **NewItemAlertRuleAction** object. If the rule has an enabled rule action that displays the specified alert in the **New item Alert** dialog box, then **[NewItemAlertRuleAction.Enabled](Outlook.NewItemAlertRuleAction.Enabled.md)** would be **True**.

For more information on specifying rule actions, see [Specify Rule Actions](http://msdn.microsoft.com/library/c5f83c81-0e01-38aa-5ec7-3932b4443e43%28Office.15%29.aspx).


## Properties



|**Name**|
|:-----|
|[ActionType](Outlook.NewItemAlertRuleAction.ActionType.md)|
|[Application](Outlook.NewItemAlertRuleAction.Application.md)|
|[Class](Outlook.NewItemAlertRuleAction.Class.md)|
|[Enabled](Outlook.NewItemAlertRuleAction.Enabled.md)|
|[Parent](Outlook.NewItemAlertRuleAction.Parent.md)|
|[Session](Outlook.NewItemAlertRuleAction.Session.md)|
|[Text](newitemalertruleaction-text-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
