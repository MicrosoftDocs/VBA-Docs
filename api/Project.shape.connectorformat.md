---
title: Shape.ConnectorFormat property (Project)
ms.prod: project-server
ms.assetid: 8bcbe86a-164e-038f-c41a-2d951e549aef
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.ConnectorFormat property (Project)
Gets a **ConnectorFormat** object that contains connector formatting properties. Applies to a **Shape** that represents a connector. Read-only **[ConnectorFormat](https://msdn.microsoft.com/library/office/ff820940%28v=office.15%29)**.

## Syntax

_expression_.**ConnectorFormat**

_expression_ A variable that represents a **[Shape](Project.Shape.md)** object.


## Remarks


> [!NOTE] 
> In Project, the connect and disconnect methods do not work for a **ConnectorFormat** object. So, the **RerouteConnections** method and the **BeginConnected**,  **BeginConnectedShape**,  **BeginConnectedSite**,  **EndConnected**,  **EndConnectedShape**, and  **EndConnectedSite** properties have no meaning.

For example, in the following code snippet, the **BeginConnect** method gives a run-time error 13, 'Type mismatch'.


```vb
Set connectorShape = oReport.Shapes.AddConnector(msoConnectorCurve, 100, 250, 150, 280)

With connectorShape
    ' Type mismatch error:
    .ConnectorFormat.BeginConnect ConnectedShape:=oReport.Shapes(5), _
        ConnectionSite:=1
    .ConnectorFormat.EndConnect ConnectedShape:=oReport.Shapes(6),_
        ConnectionSite:=1
End With
```


## Property value

 **CONNECTORFORMAT**


## See also


[Shape Object](Project.shape.md)
[AddConnector Method](Project.shapes.addconnector.md)
[ConnectorFormat](https://msdn.microsoft.com/library/office/ff820940%28v=office.15%29)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]