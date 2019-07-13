---
title: DataRecordset.Name property (Visio)
keywords: vis_sdr.chm16413930
f1_keywords:
- vis_sdr.chm16413930
ms.prod: visio
api_name:
- Visio.DataRecordset.Name
ms.assetid: 6201d472-63ee-ac51-8d08-1bf1039d8b6d
ms.date: 06/08/2017
localization_priority: Normal
---


# DataRecordset.Name property (Visio)

Gets or sets the display name of the data recordset. Read/write.


> [!NOTE] 
> This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

_expression_.**Name**

_expression_ A variable that represents a **[DataRecordset](Visio.DataRecordset.md)** object.


## Return value

String


## Remarks

The display name of a **[DataRecordset](Visio.DataRecordset.md)** object is the name that you passed as the Name parameter to the **[DataRecordsets.Add](Visio.DataRecordsets.Add.md)**, **[DataRecordsets.AddFromConnectionFile](Visio.DataRecordsets.AddFromConnectionFile.md)**, or **[DataRecordsets.AddFromXML](Visio.DataRecordsets.AddFromXML.md)** method when you first created the data recordset; or the name that you subsequently assigned by setting this property. If you do not assign a name when you create a data recordset, Microsoft Visio assigns a default name, such as _Sheet1_ , which would be the assigned name if you imported data from a Microsoft Excel workbook.

 If you specify (in the AddOptions parameter of one of the **Add** methods) that the **External Data** window be displayed in the Visio user interface, the display name appears on the tab of the **External Data** window that corresponds to the data recordset.

The **Name** property value must be unique within a particular document—you cannot assign the same display name to more than one data recordset. If you attempt to assign the **Name** property value of an existing data recordset to the another data recordset in the same document, the assignment fails. If you add a new data recordset to a document by using the **Add** method, and if you pass the **Add** method the name of an existing data recordset as a display name, Visio appends a dash and a numeral (-1, for example) to the name to make it unique within the context of the document.

However, if you delete a data recordset, you can reuse the display name of the deleted data recordset for another data recordset in the same document. This is in contrast to the case with the **[DataRecordset.ID](Visio.DataRecordset.ID.md)** property value. A particular ID is unique for the life of the document and you cannot reuse it, even if you delete the data recordset to which it was assigned.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]