---
title: ConnectorFormat.EndConnected property (Publisher)
keywords: vbapb10.chm3211523
f1_keywords:
- vbapb10.chm3211523
ms.prod: publisher
api_name:
- Publisher.ConnectorFormat.EndConnected
ms.assetid: ace997de-5a11-6b52-ac87-e914adb4212d
ms.date: 06/06/2019
localization_priority: Normal
---


# ConnectorFormat.EndConnected property (Publisher)

Returns an **[MsoTriState](Office.MsoTriState.md)** constant indicating whether the end of the specified connector is connected to a shape. Read-only.


## Syntax

_expression_.**EndConnected**

_expression_ A variable that represents a **[ConnectorFormat](Publisher.ConnectorFormat.md)** object.


## Return value

MsoTriState


## Remarks

Use the **[BeginConnected](Publisher.ConnectorFormat.BeginConnected.md)** property to determine if the beginning of a connector is connected to a shape.

The **EndConnected** property value can be one of the **MsoTriState** constants declared in the Microsoft Office type library and shown in the following table.

|Constant|Description|
|:-----|:-----|
| **msoFalse**| The end of the specified connector is not connected to a shape.|
| **msoTriStateMixed**|Return value only; indicates a combination of **msoTrue** and **msoFalse** in the specified shape range.|
| **msoTrue**| The end of the specified connector is connected to a shape.|

## Example

If the third shape on the first page in the active publication is a connector whose end is connected to a shape, this example stores the connection site number, stores a reference to the connected shape, and then disconnects the end of the connector from the shape.

```vb
Dim intSite As Integer 
Dim shpConnected As Shape 
 
With ActiveDocument.Pages(1).Shapes(3) 
 
 ' Test whether shape is a connector. 
 If .Connector Then 
 With .ConnectorFormat 
 
 ' Test whether connector is connected to another shape. 
 If .End Connected Then 
 
 ' Store connection site number. 
 intSite = .EndConnectionSite 
 
 ' Set reference to connected shape. 
 Set shpConnected = .EndConnectedShape 
 
 ' Disconnect connector and shape. 
 .EndDisconnect 
 End If 
 End With 
 End If 
End With 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]