---
title: Documents.DocumentCreated event (Visio)
keywords: vis_sdr.chm10619125
f1_keywords:
- vis_sdr.chm10619125
ms.prod: visio
api_name:
- Visio.Documents.DocumentCreated
ms.assetid: 3a436635-0148-2e24-8080-31f8e89282da
ms.date: 06/08/2017
localization_priority: Normal
---


# Documents.DocumentCreated event (Visio)

Occurs after a document is created.


## Syntax

_expression_.**DocumentCreated** (_doc_)

_expression_ A variable that represents a **[Documents](Visio.Documents.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _doc_|Required| **[IVDOCUMENT]**|The document that was created.|

## Remarks

The  **DocumentCreated** event is often added to the **EventList** collection of a Microsoft Visio template file (.vst). The event's action is triggered whenever a new document is created based on that template.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

You can add  **DocumentCreated** events to the **EventList** collection of an **Application** object, **Documents** collection, or **Document** object. The first two are straightforward; if a document is opened or created in the scope of the **Application** object or its **Documents** collection, the **DocumentCreated** event occurs.

However, adding a  **DocumentCreated** event to the **EventList** collection of a **Document** object makes sense only if the event's action is **visActCodeRunAddon**. In this case, the event is persistable; it can be stored with the document. If the document that contains the persistent event is opened, its action is triggered. If a new document is based on or copied from the document that contains the persistent event, the **DocumentCreated** event is copied to the new document and its action is triggered. However, if the event's action is **visActCodeAdvise**, that event is not persistable and therefore is not stored with the document; hence, it is never triggered.

You can prevent code from running in response to the  **DocumentCreated**, **DocumentOpened**, or **DocumentAdded** event and all events from firing by setting the value of the **EventsEnabled** property of an **Application** object to **False**.


## Example

This VBA example shows how to count shapes added to a drawing that are based on a master called  **Square**.

The  **DocumentCreated** event handler runs when a new drawing based on the template that contains this code is created. The handler initializes an integer variable, _intNumberOfSquares,_ which is used to store the count.

The  **ShapeAdded** event handler runs each time a shape is added to the drawing page, whether the shape is dragged from a stencil, drawn with a drawing tool, or pasted from the Clipboard. The handler checks the **Master** property of the new shape, and if the shape is based on the **Square** master, increments _intNumberOfSquares_.




```vb
 
Dim intNumberOfSquares As Integer 
 
Private Sub Document_DocumentCreated(ByVal vsoDocument As Visio.IVDocument) 
 
'Initialize number of squares added. 
 intNumberOfSquares = 0 
 
End Sub 
 
 
Private Sub Document_ShapeAdded(ByVal vsoShape As Visio.IVShape) 
 
 Dim vsoMaster As Visio.Master 
 
 'Get the Master property of the shape. 
 'the shape was created locally. 
 Set vsoMaster = vsoShape.Master 
 
 'Check whether the shape has a master. If not, 
 If Not (vsoMaster Is Nothing) Then 
 
 'Check whether the master is "Square". 
 If vsoMaster.Name = "Square" Then 
 
 'Increment the count for the number of squares added. 
 intNumberOfSquares = intNumberOfSquares + 1 
 
 End If 
 
 End If 
 
 MsgBox "Number of squares: " & intNumberOfSquares, vbInformation, _ 
 "Document Created Example" 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]