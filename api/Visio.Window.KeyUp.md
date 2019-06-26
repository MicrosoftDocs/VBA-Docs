---
title: Window.KeyUp event (Visio)
keywords: vis_sdr.chm11651315
f1_keywords:
- vis_sdr.chm11651315
ms.prod: visio
api_name:
- Visio.Window.KeyUp
ms.assetid: b0301a71-774b-f256-93eb-d5a3ff523def
ms.date: 06/26/2019
localization_priority: Normal
---


# Window.KeyUp event (Visio)

Occurs when a keyboard key is released.


## Syntax

_expression_.**KeyUp** (_KeyCode_, _KeyButtonState_, _CancelDefault_)

_expression_ A variable that represents a **[Window](Visio.Window.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _KeyCode_|Required| **Long**|The key that was released. Possible values are declared in [Keycode constants](../language/reference/user-interface-help/keycode-constants.md).|
| _KeyButtonState_|Required| **Long**|The state of the Shift and Ctrl keys for the event. Can be a combination of the values declared in **[VisKeyButtonFlags](visio.viskeybuttonflags.md)**. For example, if _KeyButtonState_ returns 12, it indicates that the user held down both Shift and Ctrl.|
| _CancelDefault_|Required| **Boolean**| **False** if Microsoft Visio should process the message it receives from this event; otherwise, **True**.|

## Remarks

If you set  _CancelDefault_ to **True**, Visio will not process the message received when the mouse button is clicked.

Unlike some other Visio events, **KeyUp** does not have the prefix **Query**, but it is nevertheless a query event. That is, you can cancel processing the message sent by **KeyUp**, either by setting _CancelDefault_ to **True**, or, if you are using the **[VisEventProc](visio.iviseventproc.viseventproc.md)** method to handle the event, by returning **True**. For more information, see the topics for the **VisEventProc** method and for any of the query events (for example, the **QueryCancelSuspend** event) in this reference.

If you are using Microsoft Visual Basic or VBA, the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).


## Example

This class module shows how to define a sink class called **KeyboardListener** that listens for events fired by keyboard actions in the active window. It declares the object variable _vsoWindow_ by using the **WithEvents** keyword. The class module also contains event handlers for the **KeyDown**, **KeyPress**, and **KeyUp** events.

To run this example, insert a new class module in your VBA project, name it **KeyboardListener**, and insert the following code in the module.

```vb
Dim WithEvents vsoWindow As Visio.Window 
 
Private Sub Class_Initialize() 
 
 Set vsoWindow = ActiveWindow 
 
End Sub 
 
Private Sub Class_Terminate() 
 
 Set vsoWindow = Nothing 
 
End Sub 
 
Private Sub vsoWindow_KeyDown(ByVal KeyCode As Long, ByVal KeyButtonState As Long, CancelDefault As Boolean) 
 
 Debug.Print "KeyCode is "; KeyCode 
 Debug.Print "KeyButtonState is" ; KeyButtonState 
 
End Sub 
 
Private Sub vsoWindow_KeyPress(ByVal KeyAscii As Long, CancelDefault As Boolean) 
 
 Debug.Print "KeyAscii value is "; KeyAscii 
 
End Sub 
 
Private Sub vsoWindow_KeyUp(ByVal KeyCode As Long, ByVal KeyButtonState As Long, CancelDefault As Boolean) 
 
 Debug.Print "KeyCode is "; KeyCode 
 Debug.Print "KeyButtonState is" ; KeyButtonState 
 
End Sub
```

<br/>

Then, insert the following code in the **[ThisDocument](../visio/Concepts/about-the-thisdocument-object-visio.md)** project.

```vb
Dim myKeyboardListener As KeyboardListener 
 
Private Sub Document_DocumentSaved(ByVal doc As IVDocument) 
 
 Set myKeyboardListener = New KeyboardListener 
 
End Sub 
 
Private Sub Document_BeforeDocumentClose(ByVal doc As IVDocument) 
 
 Set myKeyboardListener = Nothing 
 
End Sub
```

Save the document to initialize the class, press any key, and then release it to fire a **KeyUp** event. In the Immediate window, the handler prints the code of the key that was released to fire the event and the state of the Shift and Ctrl keys at the time the event fired.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]