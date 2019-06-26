---
title: DataRecordset.DataColumns property (Visio)
keywords: vis_sdr.chm16460285
f1_keywords:
- vis_sdr.chm16460285
ms.prod: visio
api_name:
- Visio.DataRecordset.DataColumns
ms.assetid: d22c07b9-3c92-fed4-72ed-6676ea64f1bf
ms.date: 06/08/2017
localization_priority: Normal
---


# DataRecordset.DataColumns property (Visio)

Returns the **[DataColumns](Visio.DataColumns.md)** collection associated with the **DataRecordset** object. Read-only.


> [!NOTE] 
> This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

_expression_.**DataColumns**

_expression_ An expression that returns a **[DataRecordset](Visio.DataRecordset.md)** object.


## Return value

DataColumns


## Remarks

Every **DataRecordset** object contains a **DataColumns** collection of all the **[DataColumn](Visio.DataColumn.md)** objects associated with the **DataRecordset** object. These objects allow you to map data columns to cells in the Shape Data (formerly Custom Properties) section of the Visio ShapeSheet spreadsheet.

Once you get the **DataColumns** collection, you can use its **[SetColumnProperties](Visio.DataColumns.SetColumnProperties.md)** method to set the properties of multiple data columns, or you can get and set the properties of individual data columns by using the **[DataColumn.GetProperty](Visio.DataColumn.GetProperty.md)** and **[DataColumn.SetProperty](Visio.DataColumn.SetProperty.md)** properties respectively.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]