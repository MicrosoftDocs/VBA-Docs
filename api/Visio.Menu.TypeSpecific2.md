---
title: Menu.TypeSpecific2 Property (Visio)
keywords: vis_sdr.chm13114605
f1_keywords:
- vis_sdr.chm13114605
ms.prod: visio
api_name:
- Visio.Menu.TypeSpecific2
ms.assetid: f96007e8-e459-1089-8e84-df1067d392a4
ms.date: 06/08/2017
localization_priority: Normal
---


# Menu.TypeSpecific2 Property (Visio)

Gets or sets the type of a menu. Read/write.


## Syntax

 _expression_. `TypeSpecific2`

 _expression_ A variable that represents a [Menu](./Visio.Menu.md) object.


## Return value

Integer


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

The value of an object's  **TypeSpecific2** property depends on the value of its **CntrlType** property.



|** CntrlType value**|** TypeSpecific1 value**|
|:-----|:-----|
| **visCtrlTypeBUTTON**|The  **TypeSpecific2** property is not used.|
| **visCtrlTypeCOMBOBOX**|The current width of the control expressed in pixels.|
| **visCtrlTypeEDITBOX**|The current width of the control expressed in pixels.|
| **visCtrlTypeLABEL**|The  **TypeSpecific2** property is not used.|

