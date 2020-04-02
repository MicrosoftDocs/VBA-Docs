---
title: DataRecordsets object (Visio)
keywords: vis_sdr.chm61000
f1_keywords:
- vis_sdr.chm61000
ms.prod: visio
api_name:
- Visio.DataRecordsets
ms.assetid: edf6d0dc-2f16-eee0-fd4c-ec4c9409179e
ms.date: 06/19/2019
localization_priority: Normal
---


# DataRecordsets object (Visio)

The collection of **[DataRecordset](Visio.DataRecordset.md)** objects associated with a **[Document](Visio.Document.md)** object.

> [!NOTE] 
> This Visio object or member is available only to licensed users of Visio Professional 2013.


## Remarks

The default property of the **DataRecordsets** collection is **Item**.

Every Visio **Document** object has a **DataRecordsets** collection, which is empty until you import data into Visio. To connect a Visio document to a data source, you add a **DataRecordset** object to the **DataRecordsets** collection of the document.

To add a **DataRecordset** object to the **DataRecordsets** collection, you can use one of the following three methods, depending on the type of data source that you want to connect to (OLEDB/ODBC or XML) and how you want to pass connection string and query command strings to Visio. By using the:

- **Add** method, you can connect to an OLEDB or ODBC data source and pass connection and query command string information to Visio directly as method parameters.
    
- **AddFromConnectionFile** method, you can connect to an OLEBD or ODBC data source by passing the method an Office Data Connection (ODC) file that contains the connection and query command string information that you want to supply to Visio.
    
- **AddFromXML** method, you pass the method an ADO classic XML string that contains all the data that you want to include in the data recordset.
    
After you have created a data recordset, the connection string and query command string associated with the data recordset are represented by the **[DataConnection.ConnectionString](Visio.DataConnection.ConnectionString.md)** and **[DataRecordset.CommandString](Visio.DataRecordset.CommandString.md)** properties respectively.


## Events

- [BeforeDataRecordsetDelete](Visio.DataRecordsets.BeforeDataRecordsetDelete.md)
- [DataRecordsetAdded](Visio.DataRecordsets.DataRecordsetAdded.md)
- [DataRecordsetChanged](Visio.DataRecordsets.DataRecordsetChanged.md)

## Methods

- [Add](Visio.DataRecordsets.Add.md)
- [AddFromConnectionFile](Visio.DataRecordsets.AddFromConnectionFile.md)
- [AddFromXML](Visio.DataRecordsets.AddFromXML.md)
- [GetLastDataError](Visio.DataRecordsets.GetLastDataError.md)

## Properties

- [Application](Visio.DataRecordsets.Application.md)
- [Count](Visio.DataRecordsets.Count.md)
- [Document](Visio.DataRecordsets.Document.md)
- [EventList](Visio.DataRecordsets.EventList.md)
- [Item](Visio.DataRecordsets.Item.md)
- [ItemFromID](Visio.DataRecordsets.ItemFromID.md)
- [ObjectType](Visio.DataRecordsets.ObjectType.md)
- [Stat](Visio.DataRecordsets.Stat.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]