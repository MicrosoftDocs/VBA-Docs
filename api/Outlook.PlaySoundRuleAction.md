---
title: PlaySoundRuleAction object (Outlook)
keywords: vbaol11.chm3169
f1_keywords:
- vbaol11.chm3169
ms.prod: outlook
api_name:
- Outlook.PlaySoundRuleAction
ms.assetid: 6a7a1f78-640e-8ffc-558c-c26b87638d64
ms.date: 06/08/2017
localization_priority: Normal
---


# PlaySoundRuleAction object (Outlook)

Represents an action that plays a .wav file sound.


## Remarks

 **PlaySoundRuleAction** is derived from the **[RuleAction](Outlook.RuleAction.md)** object. Each rule is associated with a **[RuleActions](Outlook.RuleActions.md)** object which has a **[PlaySound](Outlook.RuleActions.PlaySound.md)** property. The **PlaySound** property always returns a **PlaySoundRuleAction** object. If the rule has an enabled rule action that plays a sound file, then **[PlaySoundRuleAction.Enabled](Outlook.PlaySoundRuleAction.Enabled.md)** property would be **True**.

For more information on specifying rule actions, see [Specify Rule Actions](../outlook/How-to/Rules/specifying-rule-actions.md).


## Properties



|Name|
|:-----|
|[ActionType](Outlook.PlaySoundRuleAction.ActionType.md)|
|[Application](Outlook.PlaySoundRuleAction.Application.md)|
|[Class](Outlook.PlaySoundRuleAction.Class.md)|
|[Enabled](Outlook.PlaySoundRuleAction.Enabled.md)|
|[FilePath](Outlook.PlaySoundRuleAction.FilePath.md)|
|[Parent](Outlook.PlaySoundRuleAction.Parent.md)|
|[Session](Outlook.PlaySoundRuleAction.Session.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]