---
title: Styles.StyleChanged Event (Visio)
keywords: vis_sdr.chm11519250
f1_keywords:
- vis_sdr.chm11519250
ms.prod: visio
api_name:
- Visio.Styles.StyleChanged
ms.assetid: 5b9f7011-4226-e12f-a334-bd4966f4fcdf
ms.date: 06/08/2017
---


# Styles.StyleChanged Event (Visio)

Occurs after the name of a style is changed or a change to the style propagates to objects to which the style is applied.


## Syntax

Private Sub  _expression_ _'StyleChanged'(**_ByVal Style As [IVSTYLE]_**)

 _expression_ A variable that represents a [Styles](./Visio.Styles.md) object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _style_|Required| **[IVSTYLE]**|The style that changed.|

## Remarks

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).


