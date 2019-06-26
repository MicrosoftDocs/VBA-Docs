---
title: ShapeRange.ConnectorFormat property (Project)
ms.prod: project-server
ms.assetid: 7193b3aa-2e3f-d349-c398-d30e2878ceaa
ms.date: 06/08/2017
localization_priority: Normal
---


# ShapeRange.ConnectorFormat property (Project)
Gets a  **ConnectorFormat** object that contains connector formatting properties. Applies to a **ShapeRange** object that represents one or more connectors. Read-only **[ConnectorFormat](https://msdn.microsoft.com/library/office/ff820940%28v=office.15%29)**.

## Syntax

_expression_.**ConnectorFormat**

_expression_ A variable that represents a 'ShapeRange' object.


## Remarks


> [!NOTE] 
> In Project, the connect and disconnect methods do not work for a  **ConnectorFormat** object. So, the **RerouteConnections** method and the **BeginConnected**,  **BeginConnectedShape**,  **BeginConnectedSite**,  **EndConnected**,  **EndConnectedShape**, and  **EndConnectedSite** properties have no meaning.

For example, in the following code snippet, the  **BeginConnect** method gives a run-time error 13, 'Type mismatch'.


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


[ShapeRange Object](Project.shaperange.md)
[AddConnector Method](Project.shapes.addconnector.md)
[ConnectorFormat Object](https://msdn.microsoft.com/library/office/ff820940%28v=office.15%29)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]