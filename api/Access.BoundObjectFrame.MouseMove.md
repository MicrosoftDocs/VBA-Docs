---
title: BoundObjectFrame.MouseMove event (Access)
keywords: vbaac10.chm14099
f1_keywords:
- vbaac10.chm14099
ms.prod: access
api_name:
- Access.BoundObjectFrame.MouseMove
ms.assetid: 61c95d90-e3b1-a072-fcb0-66717cafb074
ms.date: 02/08/2019
localization_priority: Normal
---


# BoundObjectFrame.MouseMove event (Access)

The **MouseMove** event occurs when the user moves the mouse.


## Syntax

_expression_.**MouseMove** (_Button_, _Shift_, _X_, _Y_)

_expression_ A variable that represents a **[BoundObjectFrame](Access.BoundObjectFrame.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Button_|Required|**Integer**|The button that was pressed or released when the event was triggered. If you need to test for the _Button_ argument, you can use one of the following intrinsic constants as bit masks:<br/><br/>**acLeftButton** - The bit mask for the left mouse button.<br/>**acRightButton** - The bit mask for the right mouse button.<br/>**acMiddleButton** - The bit mask for the middle mouse button. |
| _Shift_|Required|**Integer**|The state of the Shift, Ctrl, and Alt keys when the button specified by the _Button_ argument was pressed or released. If you need to test for the _Shift_ argument, you can use one of the following intrinsic constants as bit masks:<br/><br/>**acShiftMask** - The bit mask for the Shift key.<br/>**acCtrlMask** - The bit mask for the Ctrl key.<br/>**acAltMask** - The bit mask for the Alt key.|  
| _X_|Required|**Single**|The x coordinate for the current location of the mouse pointer, in [twips](../language/glossary/vbe-glossary.md#twip). |
| _Y_|Required|**Single**|The y coordinate for the current location of the mouse pointer, in twips. |


## Remarks

- The **MouseMove** event applies only to forms, form sections, and controls on a form, not controls on a report.
    
- This event does not apply to a label attached to another control, such as the label for a text box. It applies only to "freestanding" labels. Pressing and releasing a mouse button in an attached label has the same effect as pressing and releasing the button in the associated control. The normal events for the control occur; no separate events occur for the attached label.
    


To run a macro or event procedure when these events occur, set the **OnMouseMove** property to the name of the macro or to [Event Procedure].

The **MouseMove** event is generated continually as the mouse pointer moves over objects. Unless another object generates a mouse event, an object recognizes a MouseMove event whenever the mouse pointer is positioned within its borders.

To cause a **MouseMove** event for a form to occur, move the mouse pointer over a blank area, record selector, or scroll bar on the form. To cause a **MouseMove** event for a form section to occur, move the mouse pointer over a blank area of the form section.

To respond to an event caused by moving the mouse, you use a **MouseMove** event.


 **Note**  

To run a macro or event procedure in response to pressing and releasing the mouse buttons, you use the **MouseDown** and **MouseUp** events.


## Example

The following example determines where the mouse is and whether the left mouse button and/or the SHIFT key is pressed. The x and y coordinates of the mouse pointer position are displayed in a label control as you move the mouse.


```vb
Private Sub Detail_MouseMove(Button As Integer, _ 
     Shift As Integer, X As Single, Y As Single) 
    Dim intShiftDown As Integer, intLeftButton As Integer 
 
    Me!Coordinates.Caption = X & ", " & Y 
    ' Use bit masks to determine state of 
    ' SHIFT key and left button. 
    intShiftDown = Shift And acShiftMask 
    intLeftButton = Button And acLeftButton 
    ' Check that SHIFT key and left button  
    ' are both pressed. 
    If intShiftDown And intLeftButton > 0 Then 
        MsgBox "Shift key and left mouse button were pressed." 
    End If 
End Sub
```


## See also


[BoundObjectFrame Object](Access.BoundObjectFrame.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]