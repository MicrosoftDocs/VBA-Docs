---
title: Application.MarkerEvent event (Visio)
ms.prod: visio
api_name:
- Visio.Application.MarkerEvent
ms.assetid: 1d0c20cc-ccfd-595c-04ea-afce487e582c
ms.date: 06/26/2019
localization_priority: Normal
---


# Application.MarkerEvent event (Visio)

Caused by calling the **QueueMarkerEvent** method.


## Syntax

_expression_.**MarkerEvent** (_app_, _SequenceNum_, _ContextString_)

_expression_ A variable that represents an **[Application](Visio.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _app_|Required| **[IVAPPLICATION]**|The active instance of Microsoft Visio.|
| _SequenceNum_|Required| **Long**|The ordinal position of this event with respect to past events.|
| _ContextString_|Required| **String**|Context string passed by the **QueueMarkerEvent** method.|

## Remarks

Unlike other events that Visio fires, the **MarkerEvent** event is fired by a client program. A client program receives the **MarkerEvent** event only if the client program called the **QueueMarkerEvent** method.

By using the **MarkerEvent** event in conjunction with the **QueueMarkerEvent** method, a client program can queue an event to itself. The client program receives the **MarkerEvent** event after Visio fires all the events present in its event queue at the time of the **QueueMarkerEvent** call.

The **MarkerEvent** event passes both the context string that was passed by the **QueueMarkerEvent** method and the sequence number of the **MarkerEvent** event to the **MarkerEvent** event handler. Either of these values can be used to correlate **QueueMarkerEvent** calls with **MarkerEvent** events. In this way, a client program can distinguish events it caused from those it did not cause.

For example, a client program that changes the values of Visio cells may only want to respond to the **CellChanged** events that it did not cause. The client program can first call the **QueueMarkerEvent** method and pass a context string for later use to bracket the scope of its processing.

```vb
 
vsoObject.QueueMarkerEvent "ScopeStart" 
 <My program changes cells here> 
vsoObject.QueueMarkerEvent "ScopeEnd" 

```

In the **MarkerEvent** event handler, the client program could then use the context string passed to the **QueueMarkerEvent** method to identify the **CellChanged** events that it caused.

```vb
 
Dim blsICausedCellChanges as Boolean 
 
Private Sub vsoObject_MarkerEvent (ByVal vsoApplication As Visio.IVApplication, _ 
 ByVal lngSequenceNum As Long, ByVal strContextString As String) 
 
 If strContextString = "ScopeStart" Then 
 blsICausedCellChanges = True 
 ElseIf strContextString = "ScopeEnd" Then 
 blsICausedCellChanges = "False" 
 End If 
 
End Sub 
 
Private Sub vsoObject_CellChanged (ByVal Cell As Visio.IVCell) 
 
 'Respond only if this client didn't cause a cell change. 
 If blsICausedCellChanges = False Then 
 <respond to the cell changes> 
 End If 
 
End Sub
```

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

If you are handling this event from a program that receives a notification, the **MarkerEvent** event is one of a group of events that record extra information in the **EventInfo** property of the **Application** object.

The **EventInfo** property returns _ContextString_ as described above. The _varMoreInfo_ argument to **VisEventProc** will be empty.


## Example

This example shows how to use the **MarkerEvent** event to mark an event in the event queue.

Paste this example code into the **[ThisDocument](../visio/Concepts/about-the-thisdocument-object-visio.md)** code window and then run **UseMarker**. The output is displayed in the Immediate window.

```vb
 
Dim WithEvents vsoApplication As Visio.Application 
 
Private Sub vsoApplication_MarkerEvent(ByVal app As Visio.IVApplication, _ 
 ByVal lngSequenceNum As Long, ByVal strContextString As String) 
 Debug.Print "Marker: " & app.EventInfo(0) 
 
End Sub 
 
Private Sub vsoApplication_ShapeAdded(ByVal vsoShape As Visio.IVShape) 
 Debug.Print " ShapeAdded: " & vsoShape.Name 
 
End Sub 
 
Public Sub UseMarker() 
 
 Set vsoApplication = ThisDocument.Application 
 
 'MarkerEvent events can be used to comment a segment 
 'of events in the queue. 
 vsoApplication.QueueMarkerEvent "I am starting..." 
 ActivePage.DrawRectangle 0, 0, 3, 3 
 vsoApplication.QueueMarkerEvent "I am finished..." 
 
End Sub
```

The output in the Immediate window looks like this:

> Marker: I am starting...

> ShapeAdded: Sheet.1

> Marker: I am finished...


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]