---
title: ViewCtl.ItemCount Property (Outlook View Control)
ms.prod: outlook
ms.assetid: 32c96b64-3be2-ef0b-c175-86a6f539635e
ms.date: 06/08/2017
localization_priority: Normal
---


# ViewCtl.ItemCount Property (Outlook View Control)

Returns a  **Long** that indicates the count of objects in the current folder displayed in the control. Read-only.


## Syntax

_expression_.**ItemCount**

_expression_ A variable that represents a  **ViewCtl** object.


## Remarks

The **ItemCount** property always returns the number of items that are in the current folder displayed in the control, and not the number of items that are visible in the view. Setting the [Filter](Outlook.viewctl.filt.md) or the [FilterAppend](Outlook.viewctl.filterappe.md) property has no effect on the value of the **ItemCount** property.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]