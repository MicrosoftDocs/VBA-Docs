---
title: PlaySoundRuleAction Object (Outlook)
keywords: vbaol11.chm3169
f1_keywords:
- vbaol11.chm3169
ms.prod: outlook
api_name:
- Outlook.PlaySoundRuleAction
ms.assetid: 6a7a1f78-640e-8ffc-558c-c26b87638d64
ms.date: 06/08/2017
---


# PlaySoundRuleAction Object (Outlook)

Represents an action that plays a .wav file sound.


## Remarks

 **PlaySoundRuleAction** is derived from the **[RuleAction](Outlook.RuleAction.md)** object. Each rule is associated with a **[RuleActions](Outlook.RuleActions.md)** object which has a **[PlaySound](Outlook.RuleActions.PlaySound.md)** property. The **PlaySound** property always returns a **PlaySoundRuleAction** object. If the rule has an enabled rule action that plays a sound file, then **[PlaySoundRuleAction.Enabled](Outlook.PlaySoundRuleAction.Enabled.md)** property would be **True**.

For more information on specifying rule actions, see [Specify Rule Actions](http://msdn.microsoft.com/library/c5f83c81-0e01-38aa-5ec7-3932b4443e43%28Office.15%29.aspx).


## Properties



|**Name**|
|:-----|
|[ActionType](Outlook.PlaySoundRuleAction.ActionType.md)|
|[Application](Outlook.PlaySoundRuleAction.Application.md)|
|[Class](Outlook.PlaySoundRuleAction.Class.md)|
|[Enabled](Outlook.PlaySoundRuleAction.Enabled.md)|
|[FilePath](Outlook.PlaySoundRuleAction.FilePath.md)|
|[Parent](Outlook.PlaySoundRuleAction.Parent.md)|
|[Session](playsoundruleaction-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
