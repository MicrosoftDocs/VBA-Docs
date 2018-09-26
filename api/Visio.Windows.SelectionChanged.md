---
title: Windows.SelectionChanged Event (Visio)
keywords: vis_sdr.chm11719220
f1_keywords:
- vis_sdr.chm11719220
ms.prod: visio
api_name:
- Visio.Windows.SelectionChanged
ms.assetid: 2e95eefe-5c56-8fd1-f43f-ea97602aa009
ms.date: 06/08/2017
---


# Windows.SelectionChanged Event (Visio)

Occurs after a set of shapes selected in a window changes.


## Syntax

Private Sub  _expression_ _'SelectionChanged'(**_ByVal Window As [IVWINDOW]_**)

 _expression_ A variable that represents a [Windows](./Visio.Windows.md) object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Window_|Required| **[IVWINDOW]**|The window in which the selection changed.|

## Remarks

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).


