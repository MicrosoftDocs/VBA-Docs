---
title: DataColumn.Visible property (Visio)
keywords: vis_sdr.chm16714650
f1_keywords:
- vis_sdr.chm16714650
ms.prod: visio
api_name:
- Visio.DataColumn.Visible
ms.assetid: c540f37d-abbd-4831-e43b-653b228735a2
ms.date: 06/08/2017
localization_priority: Normal
---


# DataColumn.Visible property (Visio)

Specifies whether the data column appears on the tab for its parent data recordset in the **External Data** window in the Microsoft Visio user interface. Read/write.


> [!NOTE] 
> This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

_expression_.**Visible**

_expression_ A variable that represents a **[DataColumn](Visio.DataColumn.md)** object.


## Return value

Boolean


## Remarks

If when a shape is linked to data, **Visible** is set to **True**, and if Visio adds a row to the Shape Data section of the ShapeSheet spreadsheet of the linked shape for the data column, subsequently setting the **Visible** property to **False** causes Visio to remove the ShapeSheet row it added.

If the row in the Shape Data section existed prior to linking, setting the **Visible** property to **False** does not result in Visio removing the ShapeSheet row; however, the shape data item the row represents no longer is subject to change when data in the data recordset is refreshed.

When the **Visible** property is set to **False**, Visio does not create a ShapeSheet row for the data column when it links shapes to data.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]