---
title: MenuItem.Style property (Visio)
keywords: vis_sdr.chm12951150
f1_keywords:
- vis_sdr.chm12951150
ms.prod: visio
api_name:
- Visio.MenuItem.Style
ms.assetid: 3a7cd438-2a92-b85c-5a78-2895c990f146
ms.date: 06/08/2017
localization_priority: Normal
---


# MenuItem.Style property (Visio)

Determines whether a menu item shows an icon, a caption, or some combination. Read/write.


## Syntax

_expression_.**Style**

_expression_ A variable that represents a **[MenuItem](Visio.MenuItem.md)** object.


## Return value

Integer


## Remarks


> [!NOTE] 
> Starting with Visio 2010, the Microsoft Office Fluent user interface (UI) replaced the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

Possible values for the **Style** property are listed in the following table. These constants are declared by the Visio type library in **VisUIButtonStyle**.



|Constant|Value|
|:-----|:-----|
| **visButtonAutomatic**|0|
| **visButtonCaption**|1|
| **visButtonIcon**|2|
| **visButtonIconandCaption**|3|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]