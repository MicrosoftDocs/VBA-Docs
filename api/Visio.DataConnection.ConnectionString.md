---
title: DataConnection.ConnectionString property (Visio)
keywords: vis_sdr.chm16560370
f1_keywords:
- vis_sdr.chm16560370
ms.prod: visio
api_name:
- Visio.DataConnection.ConnectionString
ms.assetid: a1a6105f-64ee-1e0c-3b54-9831aec06bf4
ms.date: 06/08/2017
localization_priority: Normal
---


# DataConnection.ConnectionString property (Visio)

Gets or sets the connection string that you can use to access an existing **[DataConnection](Visio.DataConnection.md)** object or to create a new **DataConnection** object. Read/write.


> [!NOTE] 
> This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

_expression_.**ConnectionString**

_expression_ An expression that returns a **[DataConnection](Visio.DataConnection.md)** object.


## Return value

String


## Remarks

The value of the **ConnectionString** property for a given **DataRecordset** object is the same string that you would pass to the **[DataRecordsets.Add](Visio.DataRecordsets.Add.md)** method to create the data recordset.

The easiest way to determine an appropriate connection string for a particular data source is to use the **Data Selector Wizard** in the Visio user interface (UI) to make the same connection, recording a macro while running the wizard, and then copying the connection string from the macro code.

Setting the **ConnectionString** property to a new value has no effect on data already in any data recordsets. To update the data in a data recordset using a new **ConnectionString** setting, call the **[DataRecordset.Refresh](Visio.DataRecordset.Refresh.md)** method.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]