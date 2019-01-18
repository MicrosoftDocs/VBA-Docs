---
title: DataRecordsets Object (Visio)
keywords: vis_sdr.chm61000
f1_keywords:
- vis_sdr.chm61000
ms.prod: visio
api_name:
- Visio.DataRecordsets
ms.assetid: edf6d0dc-2f16-eee0-fd4c-ec4c9409179e
ms.date: 06/08/2017
localization_priority: Normal
---


# DataRecordsets Object (Visio)

The collection of  **DataRecordset** objects associated with a **Document** object.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Remarks

The default property of the  **DataRecordsets** collection is **[Item](./Visio.DataRecordsets.Item.md)**.

Every Visio  **Document** object has a **DataRecordsets** collection, which is empty until you import data into Visio. To connect a Visio document to a data source, you add a **DataRecordset** object to the **DataRecordsets** collection of the document.

To add a  **DataRecordset** object to the **DataRecordsets** collection, you can use one of the following three methods, depending on the type of data source you want to connect to (OLEDB/ODBC or XML) and how you want to pass connection string and query command strings to Visio. By using the




-  **[DataRecordsets.Add](./Visio.DataRecordsets.Add.md)** method, you can connect to an OLEDB or ODBC data source and pass connection and query command string information to Visio directly as method parameters.
    
-  **[DataRecordsets.AddFromConnectionFile](./Visio.DataRecordsets.AddFromConnectionFile.md)** method, you can connect to an OLEBD or ODBC data source by passing the method an Office Data Connection (ODC) file that contains the connection and query command string information you want to supply to Visio.
    
-  **[DataRecordsets.AddFromXML](./Visio.DataRecordsets.AddFromXML.md)** method, you pass the method an ADO classic XML string that contains all the data that you want to include in the data recordset.
    


Once you have created a data recordset, the connection string and query command string associated with the data recordset are represented by the  **[DataConnection.ConnectionString](./Visio.DataConnection.ConnectionString.md)** and **[DataRecordset.CommandString](./Visio.DataRecordset.CommandString.md)** properties respectively.


## Events



|Name|
|:-----|
|[BeforeDataRecordsetDelete](./Visio.DataRecordset.BeforeDataRecordsetDelete.md)|
|[DataRecordsetChanged](./Visio.DataRecordset.DataRecordsetChanged.md)|

## Methods



|Name|
|:-----|
|[Delete](./Visio.DataRecordset.Delete.md)|
|[GetAllRefreshConflicts](./Visio.DataRecordset.GetAllRefreshConflicts.md)|
|[GetDataRowIDs](./Visio.DataRecordset.GetDataRowIDs.md)|
|[GetMatchingRowsForRefreshConflict](./Visio.DataRecordset.GetMatchingRowsForRefreshConflict.md)|
|[GetPrimaryKey](./Visio.DataRecordset.GetPrimaryKey.md)|
|[GetRowData](./Visio.DataRecordset.GetRowData.md)|
|[Refresh](./Visio.DataRecordset.Refresh.md)|
|[RefreshUsingXML](./Visio.DataRecordset.RefreshUsingXML.md)|
|[RemoveRefreshConflict](./Visio.DataRecordset.RemoveRefreshConflict.md)|
|[SetPrimaryKey](./Visio.DataRecordset.SetPrimaryKey.md)|

## Properties



|Name|
|:-----|
|[Application](./Visio.DataRecordset.Application.md)|
|[CommandString](./Visio.DataRecordset.CommandString.md)|
|[DataAsXML](./Visio.DataRecordset.DataAsXML.md)|
|[DataColumns](./Visio.DataRecordset.DataColumns.md)|
|[DataConnection](./Visio.DataRecordset.DataConnection.md)|
|[Document](./Visio.DataRecordset.Document.md)|
|[EventList](./Visio.DataRecordset.EventList.md)|
|[ID](./Visio.DataRecordset.ID.md)|
|[LinkReplaceBehavior](./Visio.DataRecordset.LinkReplaceBehavior.md)|
|[Name](./Visio.DataRecordset.Name.md)|
|[ObjectType](./Visio.DataRecordset.ObjectType.md)|
|[RefreshInterval](./Visio.DataRecordset.RefreshInterval.md)|
|[RefreshSettings](./Visio.DataRecordset.RefreshSettings.md)|
|[Stat](./Visio.DataRecordset.Stat.md)|
|[TimeRefreshed](./Visio.DataRecordset.TimeRefreshed.md)|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]