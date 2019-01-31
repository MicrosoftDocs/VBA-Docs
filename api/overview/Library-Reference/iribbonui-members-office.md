---
title: IRibbonUI members (Office)
description: The object that is returned by the onLoad procedure specified on the customUI tag.
ms.prod: office
ms.assetid: c6f6ec3b-3132-da29-ea08-70f20923d013
ms.date: 01/30/2019
localization_priority: Normal
---


# IRibbonUI members (Office)

The object that is returned by the **onLoad** procedure specified on the **customUI** tag. The object contains methods for invalidating control properties and for refreshing the user interface.


## Methods

|Name|Description|
|:-----|:-----|
|[ActivateTab](../../Office.IRibbonUI.ActivateTab.md)|Activates the specified custom tab. This method returns S_FALSE if there is no Ribbon or the Ribbon is collapsed.|
|[ActivateTabMso](../../Office.IRibbonUI.ActivateTabMso.md)|Activates the specified built-in tab.|
|[ActivateTabQ](../../Office.IRibbonUI.ActivateTabQ.md)|Activates the specified custom tab on the Microsoft Office Fluent Ribbon UI. Uses the fully qualified name of the tab which includes the ID and the namespace of the tab. |
|[Invalidate](../../Office.IRibbonUI.Invalidate.md)|Invalidates the cached values for all of the controls of the Ribbon user interface.|
|[InvalidateControl](../../Office.IRibbonUI.InvalidateControl.md)|Invalidates the cached value for a single control on the Ribbon user interface.|
|[InvalidateControlMso](../../Office.IRibbonUI.InvalidateControlMso.md)|Used to invalidate a built-in control.|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]