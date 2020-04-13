---
title: Shapes.AddConnector method (Project)
ms.prod: project-server
ms.assetid: bfd75cf3-f70b-8d19-bf28-94e2f4b227dd
ms.date: 06/08/2017
localization_priority: Normal
---


# Shapes.AddConnector method (Project)
Creates a connector and returns a **Shape** object the represents the new connector.

## Syntax

_expression_.**AddConnector** (_Type_, _BeginX_, _BeginY_, _EndX_, _EndY_)

_expression_ A variable that represents a **[Shapes](Project.Shapes.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Type_|Required|**MsoConnectorType**|The type of connector. Can be one of the following constants:  **msoConnectorElbow**,  **msoConnectorTypeMixed**,  **msoConnectorCurve**, or  **msoConnectorStraight**.|
| _BeginX_|Required|**Single**|The horizontal position (in points) of the connector's starting point, relative to the upper-left corner of the document.|
| _BeginY_|Required|**Single**|The vertical position (in points) of the connector's starting point.|
| _EndX_|Required|**Single**|The horizontal position (in points) of the connector's end point.|
| _EndY_|Required|**Single**|The vertical position (in points) of the connector's end point.|
| _Type_|Required|MSOCONNECTORTYPE||
| _BeginX_|Required|FLOAT||
| _BeginY_|Required|FLOAT||
| _EndX_|Required|FLOAT||
| _EndY_|Required|FLOAT||
|Name|Required/Optional|Data type|Description|

## Return value

 **Shape**


## Remarks


> [!NOTE] 
> In Project, the methods to attach the beginning and end of a connector to other shapes in the report (**ConnectorFormat.BeginConnect** and **ConnectorFormat.EndConnect**) do not work. You can use only the **AddConnector** parameters to position the connector. For more information, see the [ConnectorFormat](Project.shape.connectorformat.md) property.


## Example

The following example creates a report that contains two cloud shapes, and then adds a blue-green curved connector line that is two points wide.


```vb
Sub ConnectClouds()
    Dim shapeReport As Report
    Dim reportName As String
    Dim connectorShape As shape
    
    ' Add a report.
    reportName = "Cloud report"
    Set shapeReport = ActiveProject.Reports.Add(reportName)

    ' Add two clouds.
    Dim cloudShape1 As shape
    Dim cloudShape2 As shape
    Set cloudShape1 = shapeReport.Shapes.AddShape(msoShapeCloud, 20, 20, 100, 60)
    Set cloudShape2 = shapeReport.Shapes.AddShape(msoShapeCloud, 100, 200, 60, 100)
    
    Set connectorShape = shapeReport.Shapes.AddConnector(msoConnectorCurve, 80, 80, 130, 200)
        
    With connectorShape
        .Line.Weight = 2
        .Line.ForeColor.RGB = &HAAFF00
    End With
End Sub
```


## See also


[Shapes Object](Project.shapes.md)
[Shape Object](Project.shape.md)
[ConnectorFormat Property](Project.shape.connectorformat.md)
[AutoShapeType Property](Project.shape.autoshapetype.md)
[MsoConnectorType](https://msdn.microsoft.com/library/office/ff860918%28v=office.15%29)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]