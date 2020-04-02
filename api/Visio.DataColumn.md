---
title: DataColumn object (Visio)
keywords: vis_sdr.chm61020
f1_keywords:
- vis_sdr.chm61020
ms.prod: visio
api_name:
- Visio.DataColumn
ms.assetid: 80af7e2a-131d-515b-f582-74d903c3e02f
ms.date: 06/19/2019
localization_priority: Normal
---


# DataColumn object (Visio)

Allows custom mapping of the properties of a column of a data recordset to Microsoft Visio ShapeSheet spreadsheet cells.

> [!NOTE] 
> This Visio object or member is available only to licensed users of Visio Professional 2013.


## Remarks

The default property of the **DataColumn** object is **Name**, which returns the name of the data column in the associated data recordset.

The **Visible** property specifies whether the data column is visible in the External Data window.

Many of the properties of the **DataColumn** object correspond closely to the columns of the Shape Data section of the ShapeSheet of shapes linked to data. For example, the **DisplayName** property, which specifies the name that appears for the associated data column in the External Data window in the Visio user interface, corresponds to the Label column in the Shape Data section, which controls the label that appears for a particular Shape Data item in the **Shape Data** dialog box.

You can also set the **DisplayName** property value in the **Column Settings** dialog box in the Visio user interface (right-click in the External Data window, and then choose **Column Settings**).

> [!NOTE] 
> The read-only **Name** property specifies the programmatic name for the data column in the data recordset that contains the data column, but you can specify the value of the read/write **DisplayName** property.

> [!NOTE] 
> In Visio 2003 and prior versions, Shape Data were called Custom Properties. 

Use the **GetProperty** method to get the value of the data column property that you specify. Data column properties must be one of the enumerated values in **[VisDataColumnProperties](visio.visdatacolumnproperties.md)**, which is declared in the Visio type library.

Use the **SetProperty** method to set the value of the data column property that you specify from the members of **VisDataColumnProperties**. The **SetProperty** topic contains a table that shows a matrix of allowable data column types and property settings. These settings correspond to those you can set in the **Types and Settings** dialog box for an individual column (select a column in the **Column Settings** dialog box, and then choose **Data Type**).

## Methods

- [GetProperty](Visio.DataColumn.GetProperty.md)
- [SetProperty](Visio.DataColumn.SetProperty.md)

## Properties

- [Application](Visio.DataColumn.Application.md)
- [DataRecordset](Visio.DataColumn.DataRecordset.md)
- [DisplayName](Visio.DataColumn.DisplayName.md)
- [Document](Visio.DataColumn.Document.md)
- [Hyperlink](Visio.DataColumn.Hyperlink.md)
- [Name](Visio.DataColumn.Name.md)
- [ObjectType](Visio.DataColumn.ObjectType.md)
- [Stat](Visio.DataColumn.Stat.md)
- [Visible](Visio.DataColumn.Visible.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]