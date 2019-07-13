---
title: DataColumns object (Visio)
keywords: vis_sdr.chm61015
f1_keywords:
- vis_sdr.chm61015
ms.prod: visio
api_name:
- Visio.DataColumns
ms.assetid: 620a56f5-d552-1247-22fb-18d07993d5ad
ms.date: 06/19/2019
localization_priority: Normal
---


# DataColumns object (Visio)

The collection of **[DataColumn](Visio.DataColumn.md)** objects associated with a **[DataRecordset](visio.datarecordset.md)** object.

> [!NOTE] 
> This Visio object or member is available only to licensed users of Visio Professional 2013.


## Remarks

The default property of the **DataColumns** collection is **Item**.

A **DataRecordset** object can contain only one **DataColumns** collection. The number of **DataColumn** objects that can belong to a **DataColumns** collection is limited only by the number of columns in the data source and the hardware constraints of your computer.

You can use the **SetColumnProperties** method to set multiple properties of the data recordset columns that you specify to the values that you specify. Note that **SetColumnProperties** can set values of multiple properties for multiple columns, whereas the **[DataColumn.SetProperty](Visio.DataColumn.SetProperty.md)** method sets the value of only one property of one column at a time.

## Methods

-  [SetColumnProperties](Visio.DataColumns.SetColumnProperties.md)

## Properties

-  [Application](Visio.DataColumns.Application.md)
-  [Count](Visio.DataColumns.Count.md)
-  [DataRecordset](Visio.DataColumns.DataRecordset.md)
-  [Document](Visio.DataColumns.Document.md)
-  [Item](Visio.DataColumns.Item.md)
-  [ObjectType](Visio.DataColumns.ObjectType.md)
-  [Stat](Visio.DataColumns.Stat.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]