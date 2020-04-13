---
title: Shape.RerouteConnections method (Project)
ms.prod: project-server
ms.assetid: 97a7a245-641f-3d69-59ff-f3177ac3e84d
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.RerouteConnections method (Project)
The **RerouteConnections** method is not implemented in Project.

## Syntax

_expression_.**RerouteConnections**

_expression_ A variable that represents a **[Shape](Project.Shape.md)** object.


## Return value

 **Nothing**


## Remarks

In general for applications that implement Office Art, the **RerouteConnections** method changes the routing of a connector to the shortest path between the shapes it connects. Project does not support the connect and disconnect methods of the **ConnectorFormat** object, and so it does not support **RerouteConnections**. For more information, see the **[ConnectorFormat](Project.shape.connectorformat.md)** property.


## See also


[Shape Object](Project.shape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]