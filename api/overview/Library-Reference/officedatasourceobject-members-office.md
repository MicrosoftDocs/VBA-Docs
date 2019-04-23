---
title: OfficeDataSourceObject members (Office)
ms.prod: office
ms.assetid: 57ba0dc6-80e7-04a9-a619-2a3e6aa2cdff
ms.date: 01/30/2019
localization_priority: Normal
---


# OfficeDataSourceObject members (Office)

Represents the mail merge data source in a mail merge operation.


## Methods

|Name|Description|
|:-----|:-----|
|[ApplyFilter](../../Office.OfficeDataSourceObject.ApplyFilter.md)|Applies a filter to a mail merge data source to filter specified records meeting specified criteria.|
|[Move](../../Office.OfficeDataSourceObject.Move.md)|Moves a record in a return set from an **OfficeDataSourceObject** object from one position to another.|
|[Open](../../Office.OfficeDataSourceObject.Open.md)|Opens a table in an **OfficeDataSourceObject** object.|
|[SetSortOrder](../../Office.OfficeDataSourceObject.SetSortOrder.md)|Sets the sort order for mail merge data.|


## Properties

|Name|Description|
|:-----|:-----|
|[Columns](../../Office.OfficeDataSourceObject.Columns.md)|Gets an **ODSOColumns** object that represents the fields in a data source. Read-only.|
|[ConnectString](../../Office.OfficeDataSourceObject.ConnectString.md)|Gets or sets a **String** that represents the connection to the specified mail merge data source. Read/write.|
|[DataSource](../../Office.OfficeDataSourceObject.DataSource.md)|Gets or sets a **String** that represents the name of the attached data source. Read/write.|
|[Filters](../../Office.OfficeDataSourceObject.Filters.md)|Gets the filter status for a **OfficeDataSourceObject** object. Read-only.|
|[RowCount](../../Office.OfficeDataSourceObject.RowCount.md)|Gets a **Long** that represents the number of records in the specified data source. Read-only.|
|[Table](../../Office.OfficeDataSourceObject.Table.md)|Gets a **String** that represents the name of the table within the data source file that contains the mail merge records. The returned value may be blank if the table name is unknown or not applicable to the current data source. Read-only.|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]