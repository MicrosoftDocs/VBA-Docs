---
title: DataRecordset.Delete method (Visio)
keywords: vis_sdr.chm16416165
f1_keywords:
- vis_sdr.chm16416165
ms.prod: visio
api_name:
- Visio.DataRecordset.Delete
ms.assetid: 9f3fa9b0-2ca9-cf28-fa27-18eef4be179d
ms.date: 06/08/2017
localization_priority: Normal
---


# DataRecordset.Delete method (Visio)

Deletes the **[DataRecordset](Visio.DataRecordset.md)** object from the **[DataRecordsets](Visio.DataRecordsets.md)** collection of the document. .


> [!NOTE] 
> This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

_expression_.**Delete**

_expression_ A variable that represents a **[DataRecordset](Visio.DataRecordset.md)** object.


## Return value

Nothing


## Remarks

If the **DataRecordset** object to be deleted is associated with a **[DataConnection](Visio.DataConnection.md)** object, and if that **DataConnection** object is not associated with any other **DataRecordset** objects, Microsoft Visio also deletes the **DataConnection** object.

Note that deleting a **DataRecordset** object does not delete the shapes that had been linked to data in that data recordset, nor does delete any existing shape data in those shapes that was created when the shapes were linked to data.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the **Delete** method to delete a **DataRecordset** from the **DataRecordsets** collection of the current document. It gets the count of all data recordsets associated with the current document and deletes the one most recently added.


```vb
Public Sub Delete_Example() 
 
    Dim vsoDataRecordset As Visio.DataRecordset 
    Dim intCount As Integer 
 
    intCount = ThisDocument.DataRecordsets.Count 
    Set vsoDataRecordset = ThisDocument.DataRecordsets(intCount) 
    vsoDataRecordset.Delete 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]