---
title: AfterUpdate Event
keywords: fm20.chm5224934
f1_keywords:
- fm20.chm5224934
ms.prod: office
api_name:
- Office.AfterUpdate
ms.assetid: 3d15efd4-06c8-136f-c315-7efc44db35b1
ms.date: 06/08/2017
---


# AfterUpdate Event



Occurs after data in a control is changed through the user interface.
<<<<<<< HEAD
 **Syntax**
 **Private Sub**_object_ _**AfterUpdate( )**
=======

## Syntax

**Private Sub**_object_ _**AfterUpdate( )**
>>>>>>> master
The  **AfterUpdate** event syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|

<<<<<<< HEAD
 **Remarks**
The AfterUpdate event occurs regardless of whether the control is [bound](../../../language/Glossary/glossary-vba.md) (that is, when the **RowSource** property specifies a[data source](../../../language/Glossary/glossary-vba.md) for the control). This event cannot be canceled. If you want to cancel the update (to restore the previous value of the control), use the BeforeUpdate event and set the _Cancel_ argument to **True**.
The AfterUpdate event occurs after the BeforeUpdate event and before the Exit event for the current control and before the Enter event for the next control in the [tab order](../../Glossary/vbe-glossary.md).
=======
## Remarks

The AfterUpdate event occurs regardless of whether the control is [bound](../../Glossary/glossary-vba.md#bound) (that is, when the **RowSource** property specifies a[data source](../../Glossary/glossary-vba.md#data-source) for the control). This event cannot be canceled. If you want to cancel the update (to restore the previous value of the control), use the BeforeUpdate event and set the _Cancel_ argument to **True**.
The AfterUpdate event occurs after the BeforeUpdate event and before the Exit event for the current control and before the Enter event for the next control in the [tab order](../../Glossary/vbe-glossary.md#tab-order).
>>>>>>> master

