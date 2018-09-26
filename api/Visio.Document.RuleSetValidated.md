---
title: Document.RuleSetValidated Event (Visio)
keywords: vis_sdr.chm10562085
f1_keywords:
- vis_sdr.chm10562085
ms.prod: visio
api_name:
- Visio.Document.RuleSetValidated
ms.assetid: 682b8f48-4ebe-ce53-f816-3d82a4ae0034
ms.date: 06/08/2017
---


# Document.RuleSetValidated Event (Visio)

Occurs when a rule set is validated.


## Syntax

Private Sub  _expression_ _'RuleSetValidated'(**_ByVal RuleSet As ValidationRuleSet_**)

 _expression_ A variable that represents a '[Document](Visio.Document.md)' object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _RuleSet_|Required| **[ValidationRuleSet](Visio.ValidationRuleSet.md)**|The rule set that was validated.|

## Remarks

When Microsoft Visio performs validation, it fires a  **RuleSetValidated** event for every rule set that it processes, even if a rule set is empty.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **[Event](Visio.Event.md)** objects, use the **[EventList.Add](Visio.EventList.Add.md)** or **[EventList.AddAdvise](Visio.EventList.AddAdvise.md)** method. To create an **Event** object that runs an add-on, use the **EventList.Add** method. To create an **Event** object that receives notification, use the **EventList.AddAdvise** method. To find an event code for the event you want to create, see [Event Codes](../visio/Concepts/event-codesvisio.md).


