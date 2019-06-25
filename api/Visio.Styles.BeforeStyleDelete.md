---
title: Styles.BeforeStyleDelete event (Visio)
keywords: vis_sdr.chm11519070
f1_keywords:
- vis_sdr.chm11519070
ms.prod: visio
api_name:
- Visio.Styles.BeforeStyleDelete
ms.assetid: e73533d6-c5ce-739c-f85d-0137794ac953
ms.date: 06/08/2017
localization_priority: Normal
---


# Styles.BeforeStyleDelete event (Visio)

Occurs before a style is deleted.


## Syntax

_expression_.**BeforeStyleDelete** (_Style_)

_expression_ A variable that represents a **[Styles](Visio.Styles.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Style_|Required| **[IVSTYLE]**|The style that is going to be deleted.|

## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]