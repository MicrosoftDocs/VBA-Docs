---
title: Document.DocumentSaved Event (Visio)
keywords: vis_sdr.chm10519135
f1_keywords:
- vis_sdr.chm10519135
ms.prod: visio
api_name:
- Visio.Document.DocumentSaved
ms.assetid: 48e513a1-4382-eb3c-4838-ad2f85483f51
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.DocumentSaved Event (Visio)

Occurs after a document is saved.


## Syntax

Private Sub  _expression_ _'DocumentSaved'(**_ByVal doc As [IVDOCUMENT]_**)

 _expression_ A variable that represents a [Document](./Visio.Document.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _doc_|Required| **[IVDOCUMENT]**|The document that was saved.|

## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).


## Example

This example shows how to create a class module to handle events fired by a source object in Microsoft Visio, for example, the  **Document** object. The module consists of the function **VisEventProc** , which uses a **Select Case** block to check for three events: **DocumentSaved** , **PageAdded** , and **ShapesDeleted**. Other events fall under the default case (**Case Else**). Each **Case** block constructs a string ( _strMessage_ ) that contains the name and event code of the event that fired. Finally, the function displays the string in the Immediate window.

Copy this sample code into a new class module in VBA or Visual Basic, naming the module  **clsEventSink**. You can then use an event-sink module to create an instance of the **clsEventSink** class and **Event** objects that send notifications of event firings to the class instance. To see how to create an event-sink module, refer to the example for the **AddAdvise** method.




```vb
 
Implements Visio.IVisEventProc 
 
'Declare visEvtAdd as a 2-byte value 
'to avoid a run-time overflow error 
Private Const visEvtAdd% = &H8000 
 
Private Function IVisEventProc_VisEventProc( _ 
 ByVal nEventCode As Integer, _ 
 ByVal pSourceObj As Object, _ 
 ByVal nEventID As Long, _ 
 ByVal nEventSeqNum As Long, _ 
 ByVal pSubjectObj As Object, _ 
 ByVal vMoreInfo As Variant) As Variant 
 
 Dim strMessage As String 
 
 'Find out which event fired 
 Select Case nEventCode 
 Case visEvtCodeDocSave 
 strMessage = "DocumentSaved (" & nEventCode & ")" 
 Case (visEvtPage + visEvtAdd) 
 strMessage = "PageAdded (" & nEventCode & ")" 
 Case visEvtCodeShapeDelete 
 strMessage = "ShapesDeleted(" & nEventCode & ")" 
 Case Else 
 strMessage = "Other (" & nEventCode & ")" 
 End Select 
 
 'Display the event name and the event code 
 Debug.Print strMessage 
 
End Function
```


