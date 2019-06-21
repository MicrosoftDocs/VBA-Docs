---
title: Page.GetShapesLinkedToData method (Visio)
keywords: vis_sdr.chm10960145
f1_keywords:
- vis_sdr.chm10960145
ms.prod: visio
api_name:
- Visio.Page.GetShapesLinkedToData
ms.assetid: 3196f7f9-1b7c-8070-444d-c1a55f0c205f
ms.date: 06/08/2017
localization_priority: Normal
---


# Page.GetShapesLinkedToData method (Visio)

Returns an array of all shapes on the active page linked to data in the specified data recordset.


> [!NOTE] 
> This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

_expression_. `GetShapesLinkedToData`( `_DataRecordsetID_` , `_ShapeIDs()_` )

 _expression_ An expression that returns a **[Page](Visio.Page.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _DataRecordsetID_|Required| **Long**|The ID of a data recordset contained in the current document.|
| _ShapeIDs()_|Required| **Long**|Out parameter. An array of type  **Long** that the method will return filled with the shape IDs of shapes on the page linked to the data recordset specified in DataRecordsetID.|

## Return value

Nothing


## Remarks

For the ShapeIDs() parameter, pass an empty, dimensionless array of type  **Long**. If there are no linked shapes on the page, **GetShapesLinkedToData** will return an empty array.

To determine the specific data row in the data recordset shapes are linked to, use the  **[Page.GetShapesLinkedToDataRow](Visio.Page.GetShapesLinkedToDataRow.md)** method.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **GetShapesLinkedToData** method to determine the shape IDs of the shapes on the page linked to data in the data recordset most recently added to the **DataRecordsets** collection of the current document. It prints the shape IDs in the Immediate window.

Before running this macro, use the  **[DataRecordsets.Add](Visio.DataRecordsets.Add.md)** method or another means to add at least one data recordset to the **DataRecordsets** collection, and make sure there is at least one shape on the active page linked to data in the data recordset.




```vb
Public Sub GetShapesLinkedToData_Example() 
 
    Dim vsoDataRecordset As Visio.DataRecordset 
    Dim intRecordsetCount As Integer 
    Dim alngShapeIDs() As Long 
    Dim intArrayCounter As Integer 
     
    intRecordsetCount = Visio.ActiveDocument.DataRecordsets.Count 
    Set vsoDataRecordset = Visio.ActiveDocument.DataRecordsets(intRecordsetCount) 
     
    ActivePage.GetShapesLinkedToData vsoDataRecordset.ID, alngShapeIDs 
     
    For intArrayCounter = LBound(alngShapeIDs) To UBound(alngShapeIDs) 
        Debug.Print alngShapeIDs(intArrayCounter) 
    Next 
     
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]