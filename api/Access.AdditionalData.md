---
title: AdditionalData object (Access)
keywords: vbaac10.chm13253
f1_keywords:
- vbaac10.chm13253
ms.prod: access
api_name:
- Access.AdditionalData
ms.assetid: 2677072b-c2ca-3bcd-fef4-f6b1cadb0379
ms.date: 02/01/2019
localization_priority: Normal
---


# AdditionalData object (Access)

Represents the collection of tables and queries that will be included with the parent table that is exported by the **[ExportXML](Access.Application.ExportXML.md)** method.


## Remarks

To create an **AdditionalData** object, use the **[CreateAdditionalData](Access.Application.CreateAdditionalData.md)** method of the **[Application](Access.Application.md)** object.

To add a table to an existing **AdditionalData** object, use the **Add** method.


## Example

The following example exports the contents of the Customers table in the Northwind Traders sample database, along with the contents of the Orders and Orders Details tables, to an XML data file named Customer Orders.xml.


```vb
Sub ExportCustomerOrderData() 
 Dim objOrderInfo As AdditionalData 
 
 Set objOrderInfo = Application.CreateAdditionalData 
 
 ' Add the Orders and Order Details tables to the data to be exported. 
 objOrderInfo.Add "Orders" 
 objOrderInfo.Add "Order Details" 
 
 ' Export the contents of the Customers table. The Orders and Order 
 ' Details tables will be included in the XML file. 
 Application.ExportXML ObjectType:=acExportTable, DataSource:="Customers", _ 
 DataTarget:="Customer Orders.xml", _ 
 AdditionalData:=objOrderInfo 
End Sub
```


## Methods

- [Add](Access.AdditionalData.Add.md)

## Properties

- [Count](Access.AdditionalData.Count.md)
- [Item](Access.AdditionalData.Item.md)
- [Name](Access.AdditionalData.Name.md)

## See also

- [Access Object Model Reference](overview/Access/object-model.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]