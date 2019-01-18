---
title: ZOrder method (Microsoft Forms)
keywords: fm20.chm5224976
f1_keywords:
- fm20.chm5224976
ms.prod: office
api_name:
- Office.ZOrder
ms.assetid: dcf6f2b8-9f00-a8a7-2911-bfee9027a6f3
ms.date: 11/15/2018
localization_priority: Normal
---


# ZOrder method (Microsoft Forms)

Places the object at the front or back of the [z-order](../../Glossary/vbe-glossary.md#z-order).

## Syntax

_object_. **ZOrder(** [ _zPosition_ ] **)**

The **ZOrder** method syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|
| _zPosition_|Optional. A control's position, front or back, in the container's z-order.|

## Settings

The settings for _zPosition_ are:

|Constant|Value|Description|
|:-----|:-----|:-----|
| _fmTop_|0|Places the control at the front of the z-order. The control appears on top of other controls (default).|
| _fmBottom_|1|Places the control at the back of the z-order. The control appears underneath other controls.|

## Remarks

The z-order determines how windows and controls are stacked when they are presented to the user. Items at the back of the z-order are overlaid by closer items; items at the front of the z-order appear to be on top of items at the back. When the _zPosition_ argument is omitted, the object is brought to the front.

In [design mode](../../Glossary/vbe-glossary.md#document-design-mode), the **Bring to Front** or **Send to Back** commands set the z-order. **Bring to Front** is equivalent to using the **ZOrder** method and putting the object at the front of the z-order. **Send to Back** is equivalent to using **ZOrder** and putting the object at the back of the z-order.

This method does not affect content or sequence of the controls in the **Controls** collection.

> [!NOTE] 
> You can't undo or redo layering commands, such as **Send to Back** or **Bring to Front**. For example, if you select an object and click **Move Backward** on the shortcut menu, you won't be able to undo or redo that action.


## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]