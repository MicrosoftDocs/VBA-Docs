---
title: Styles.StyleAdded Event (Visio)
keywords: vis_sdr.chm11519245
f1_keywords:
- vis_sdr.chm11519245
ms.prod: visio
api_name:
- Visio.Styles.StyleAdded
ms.assetid: c5bee31d-1b73-dd43-c44a-094004abd2ea
ms.date: 06/08/2017
---


# Styles.StyleAdded Event (Visio)

Occurs after a new style is added to a document.


## Syntax

Private Sub  _expression_ _'StyleAdded'(**_ByVal Style As [IVSTYLE]_**)

 _expression_ A variable that represents a [Styles](./Visio.Styles.md) object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Style_|Required| **[IVSTYLE]**|The style that was added to the document.|

## Remarks

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).


