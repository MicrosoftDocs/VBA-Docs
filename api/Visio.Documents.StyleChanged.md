---
title: Documents.StyleChanged event (Visio)
keywords: vis_sdr.chm10619250
f1_keywords:
- vis_sdr.chm10619250
ms.prod: visio
api_name:
- Visio.Documents.StyleChanged
ms.assetid: 2ec52d84-fc79-4798-f01d-bc594ca39bd7
ms.date: 06/08/2017
localization_priority: Normal
---


# Documents.StyleChanged event (Visio)

Occurs after the name of a style is changed or a change to the style propagates to objects to which the style is applied.


## Syntax

_expression_.**StyleChanged** (_Style_)

_expression_ A variable that represents a **[Documents](Visio.Documents.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Style_|Required| **[IVSTYLE]**|The style that changed.|

## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]