---
title: DataRecordset.DataConnection property (Visio)
keywords: vis_sdr.chm16460280
f1_keywords:
- vis_sdr.chm16460280
ms.prod: visio
api_name:
- Visio.DataRecordset.DataConnection
ms.assetid: 3425e9c4-4cd6-7553-2dbf-5e14b8a9a68a
ms.date: 06/08/2017
localization_priority: Normal
---


# DataRecordset.DataConnection property (Visio)

Returns the **[DataConnection](Visio.DataConnection.md)** object associated with the **DataRecordset** object. Read-only.


> [!NOTE] 
> This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

_expression_.**DataConnection**

_expression_ An expression that returns a **[DataRecordset](Visio.DataRecordset.md)** object.


## Return value

DataConnection


## Remarks

You can get the connection string associated with a data recordset by first using the **DataConnection** property to get the **DataConnection** object associated with the data recordset and then getting the **[DataConnection.ConnectionString](Visio.DataConnection.ConnectionString.md)** property value.

The **DataConnection** property returns **Nothing** for "connectionless" **DataRecordset** objects—those that are created by using the **[DataRecordsets.AddFromXML](Visio.DataRecordsets.AddFromXML.md)** method.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]