---
title: UIObject.MenuAnimationStyle property (Visio)
keywords: vis_sdr.chm14913900
f1_keywords:
- vis_sdr.chm14913900
ms.prod: visio
api_name:
- Visio.UIObject.MenuAnimationStyle
ms.assetid: 17a7b713-62b4-98cc-141d-fd86e762ba99
ms.date: 06/08/2017
localization_priority: Normal
---


# UIObject.MenuAnimationStyle property (Visio)

Gets or sets the way in which a menu is displayed. Read/write.


## Syntax

_expression_. `MenuAnimationStyle`

_expression_ A variable that represents a **[UIObject](Visio.UIObject.md)** object.


## Return value

Integer


## Remarks

> [!NOTE] 
> Starting with Visio 2010, the Microsoft Office Fluent user interface (UI) replaced the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

You can use any **UIObject** object to get or set this property. The property affects the entire application and affects the appearance of buttons in the currently visible set of toolbars.

Constants representing animation styles are prefixed with **visMenuAnimation** and are declared by the Visio type library in member **VisUIMenuAnimation**.

|Constant|Value|
|:-----|:-----|
| **visMenuAnimationNone**| 0|
| **visMenuAnimationRandom**| 1|
| **visMenuAnimationUnfold**| 2|
| **visMenuAnimationSlide**| 3|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]