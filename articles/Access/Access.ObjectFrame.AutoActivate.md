---
title: ObjectFrame.AutoActivate Property (Access)
keywords: vbaac10.chm11567
f1_keywords:
- vbaac10.chm11567
ms.prod: access
api_name:
- Access.ObjectFrame.AutoActivate
ms.assetid: e6e0dfce-1bfe-707b-d7f0-45a216d4aa55
ms.date: 06/08/2017
---


# ObjectFrame.AutoActivate Property (Access)

You can use the  **AutoActivate** property to specify how the user can activate an OLE object. Read/write **Integer**.


## Syntax

 _expression_. **AutoActivate**

 _expression_ A variable that represents an **ObjectFrame** object.


## Remarks

The  **AutoActivate** property uses the following settings.



|**Setting**|**Constant**|**Description**|
|:-----|:-----|:-----|
|Manual|**acOLEActivateManual** (0)|The OLE object isn't activated when it receives the focus or when the user double-clicks the control. You can activate an OLE object only by using Visual Basic to set the control's  **Action** property to **acOLEActivate**.|
|GetFocus|**acOLEActivateGetFocus** (1)|(For unbound object frame and chart controls only) If the control contains an OLE object, the application that supplied the object is activated when the control receives the focus.|
|Double-Click|**acOLEActivateDoubleClick** (2)|(Default) If the control contains an OLE object, the application that supplied the object is activated when the user double-clicks the control or presses CTRL+ENTER when the control has the focus.|
The  **AutoActivate** property can be set only in Design view.

Some OLE objects can be activated from within the control. When such an object is activated, the object can be edited (or some other operation can be performed) from inside the boundaries of the control. This feature is called in-place activation. If an object supports in-place activation, see the documentation for the application that was used to create the object for information about using this feature.

With Visual Basic, you can determine if a control contains an object by checking the setting of its  **OLEType** property.


 **Note**   If you set a control's **AutoActivate** property to Double-Click and specify a **DblClick** event for the control, the DblClick event occurs before the object is activated.


## See also


#### Concepts


[ObjectFrame Object](Access.ObjectFrame.md)

