---
title: Shape.GetLinkedDataRecordsetIDs method (Visio)
keywords: vis_sdr.chm11260220
f1_keywords:
- vis_sdr.chm11260220
ms.prod: visio
api_name:
- Visio.Shape.GetLinkedDataRecordsetIDs
ms.assetid: 1ce55d6c-02ae-8d5d-f581-b368e830bcf5
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.GetLinkedDataRecordsetIDs method (Visio)

Gets the IDs of all the data recordsets that contain data rows linked to the shape.


> [!NOTE] 
> This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

_expression_. `GetLinkedDataRecordsetIDs`( `_DataRecordsetIDs()_` )

 _expression_ An expression that returns a **[Shape](Visio.Shape.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _DataRecordsetIDs()_|Required| **Long**|Out parameter. An array of IDs of data recordsets containing data rows linked to the shape.|

## Return value

Nothing


## Remarks

For the DataRecordsetIDs() parameter, pass an empty, dimensionless array of type  **Long** that the method fills with the IDs of data recordsets containing data rows linked to the shape.


## Example

 The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the **GetLinkedDataRecordsetIDs** method to get the IDs of all the data recordsets that contain data rows linked to the shape.

Before running this macro, add at least two data recordsets to the  **[DataRecordsets](Visio.DataRecordsets.md)** collection of the document. The macro drops a shape onto the page, links the shape to data in the two data recordsets most recently added to the collection, and then prints the IDs of the linked data recordsets returned by the **GetLinkedDataRecordsetIDs** method in the Immediate window.




```vb
Public Sub GetLinkedDataRecordsetIDs_Example() 
 
    Dim vsoDataRecordset1 As Visio.DataRecordset 
    Dim vsoDataRecordset2 As Visio.DataRecordset 
    Dim vsoShape As Visio.Shape 
    Dim intCount As Integer 
    Dim alngDataRecordsetIDs() As Long 
    Dim intArrayIndex As Integer 
     
    intCount = Visio.ActiveDocument.DataRecordsets.Count 
    Set vsoDataRecordset1 = Visio.ActiveDocument.DataRecordsets(intCount) 
    Set vsoDataRecordset2 = Visio.ActiveDocument.DataRecordsets(intCount - 1) 
     
    Set vsoShape = ActivePage.DrawRectangle(2, 2, 4, 4) 
         
    vsoShape.LinkToData vsoDataRecordset1.ID, 1, True 
    vsoShape.LinkToData vsoDataRecordset2.ID, 2, True 
         
    vsoShape.GetLinkedDataRecordsetIDs alngDataRecordsetIDs 
         
    For intArrayIndex = LBound(alngDataRecordsetIDs) To UBound(alngDataRecordsetIDs) 
        Debug.Print alngDataRecordsetIDs(intArrayIndex) 
    Next 
         
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]