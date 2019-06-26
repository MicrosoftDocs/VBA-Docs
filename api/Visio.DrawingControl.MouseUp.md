---
title: DrawingControl.MouseUp event (Visio)
ms.prod: visio
api_name:
- Visio.DrawingControl.MouseUp
ms.assetid: 34f7d931-5f4d-523e-b4d8-9096c4a634c3
ms.date: 06/26/2019
localization_priority: Normal
---


# DrawingControl.MouseUp event (Visio)

Occurs when a mouse button is released.


## Syntax

_expression_.**MouseUp** (_Button_, _KeyButtonState_, _x_, _y_, _CancelDefault_)

_expression_ A variable that represents a **[DrawingControl](Visio.DrawingControl.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Button_|Required| **Long**|The mouse button that was released. Possible values are declared in **[VisKeyButtonFlags](visio.viskeybuttonflags.md)**.|
| _KeyButtonState_|Required| **Long**|The state of the mouse buttons and the Shift and Ctrl keys for the event. Possible values can be a combination of the values declared in **VisKeyButtonFlags**. For example, if _KeyButtonState_ returns 9, it indicates that the user clicked the left mouse button while pressing Ctrl.|
| _x_|Required| **Double**|The x-coordinate of the mouse pointer.|
| _y_|Required| **Double**|The y-coordinate of the mouse pointer.|
| _CancelDefault_|Required| **Boolean**| **False** if Microsoft Visio should process the message it receives from this event; otherwise, **True**.|

## Remarks

If you set _CancelDefault_ to **True**, Visio will not process the message received when the mouse button is clicked.

Unlike some other Visio events, **MouseUp** does not have the prefix **Query**, but it is nevertheless a query event. That is, you can cancel processing the message sent by **MouseUp**, either by setting _CancelDefault_ to **True**, or, if you are using the **[VisEventProc](visio.iviseventproc.viseventproc.md)** method to handle the event, by returning **True**. For more information, see the topics for the **VisEventProc** method and for any of the query events (for example, the **QueryCancelSuspend** event) in this reference.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).


## Example

This class module shows how to define a sink class called **MouseListener** that listens for events fired by mouse actions in the active window. It declares the object variable _vsoWindow_ by using the **WithEvents** keyword. The class module also contains event handlers for the **MouseDown**, **MouseMove**, and **MouseUp** events.

To run this example, insert a new class module in your VBA project, name it **MouseListener**, and insert the following code in the module.

```vb
Dim WithEvents vsoWindow As Visio.Window 
 
Private Sub Class_Initialize() 
 
 Set vsoWindow = ActiveWindow 
 
End Sub 
 
Private Sub Class_Terminate() 
 
 Set vsoWindow = Nothing 
 
End Sub 
 
Private Sub vsoWindow_MouseDown(ByVal Button As Long, ByVal KeyButtonState As Long, ByVal x As Double, ByVal y As Double, CancelDefault As Boolean) 
 
 If Button = 1 Then 
 
 Debug.Print "Left mouse button clicked" 
 
 ElseIf Button = 2 Then 
 
 Debug.Print "Right mouse button clicked" 
 
 ElseIf Button = 16 Then 
 
 Debug.Print "Center mouse button clicked" 
 
 End If 
 
End Sub 
 
Private Sub vsoWindow_MouseMove(ByVal Button As Long, ByVal KeyButtonState As Long, ByVal x As Double, ByVal y As Double, CancelDefault As Boolean) 
 
 Debug.Print "x-position is "; x 
 Debug.Print "y-position is "; y 
 
End Sub 
 
Private Sub vsoWindow_MouseUp(ByVal Button As Long, ByVal KeyButtonState As Long, ByVal x As Double, ByVal y As Double, CancelDefault As Boolean) 
 
 If Button = 1 Then 
 
 Debug.Print "Left mouse button released" 
 
 ElseIf Button = 2 Then 
 
 Debug.Print "Right mouse button released" 
 
 ElseIf Button = 16 Then 
 
 Debug.Print "Center mouse button released" 
 
 End If 
 
End Sub
```

<br/>

Then, insert the following code in the **[ThisDocument](../visio/Concepts/about-the-thisdocument-object-visio.md)** project.

```vb
Dim myMouseListener As MouseListener 
 
Private Sub Document_DocumentSaved(ByVal doc As IVDocument) 
 
 Set myMouseListener = New MouseListener 
 
End Sub 
 
Private Sub Document_BeforeDocumentClose(ByVal doc As IVDocument) 
 
 Set myMouseListener = Nothing 
 
End Sub
```

Save the document to initialize the class, and then click anywhere in the active window to fire a **MouseUp** event. In the Immediate window, the handler prints the name of the mouse button that was clicked to fire the event.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]