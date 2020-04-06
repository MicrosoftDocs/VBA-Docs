---
title: Shapes.Range method (Project)
ms.prod: project-server
ms.assetid: 984326ae-f567-18b8-562a-fcb2160b0dad
ms.date: 06/08/2017
localization_priority: Normal
---


# Shapes.Range method (Project)
Returns a  **ShapeRange** object that represents a subset of shapes in the **Shapes** collection.

## Syntax

_expression_.**Range** (_Index_)

_expression_ A variable that represents a **[Shapes](Project.Shapes.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Variant**|Specifies one or more shapes to be included in the range. Can be an integer for the index number of a shape, a string for the name of a shape, or an array that contains either integers or strings.|
| _Index_|Required|**Variant**||
|Name|Required/Optional|Data type|Description|

## Return value

 **ShapeRange**

The range of shapes that are specified by the  _Index_ parameter.


## Remarks


> [!NOTE] 
> Most operations that you can do with a  **Shape** object you can also do with a **ShapeRange** object that contains a single shape. Some operations, when performed on a **ShapeRange** object that contains multiple shapes, produce an error.

Although you can use the  **Range** property to return any number of shapes on a report, it is simpler to use the default **Value** property to return a single **Shape** in the collection. For example, `Shapes(1)` is simpler than `Shapes.Range(1)`.

To specify an array of integers or strings for the  _Index_ parameter, you can use the **Array** function. For example, the following macro selects two shapes that are specified by name.




```vb
Sub SelectShapeRange()
    Dim arShapes() As Variant
    Dim oShapeRange As ShapeRange
    
    arShapes = Array("TextBox 4", "TextBox 5")
    Set oShapeRange = ActiveProject.Reports("Table Tests").Shapes.Range(arShapes)
    oShapeRange.Select
End Sub
```


## Example

If you create a report that has two text boxes such as in the previous code, the following macro selects the text boxes by index number, and then adds a shadow to each of them.


```vb
Sub AddShadow2Shapes()
    Dim oReports As Reports
    Dim oReport As Report
    Dim oShapeRange As ShapeRange
    Dim reportName As String
    Dim arShapes() As Variant

    arShapes = Array(3, 4)

    reportName = "Table Tests"
    Set oReports = ActiveProject.Reports
    
    If (oReports.IsPresent(reportName)) Then
        ' Make the report the active view.
        oReports(reportName).Apply
        
        Set oReport = oReports(reportName)
        
        Set oShapeRange = oReport.Shapes.Range(arShapes)
        
        oShapeRange.Select
        oShapeRange.Shadow.Type = msoShadow1
    End If
End Sub
```


## See also


[Shapes Object](Project.shapes.md)
[ShapeRange Object](Project.shaperange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]