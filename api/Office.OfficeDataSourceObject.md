---
title: OfficeDataSourceObject Object (Office)
keywords: vbaof11.chm232000
f1_keywords:
- vbaof11.chm232000
ms.prod: office
api_name:
- Office.OfficeDataSourceObject
ms.assetid: d5e5401b-643e-c12c-2648-f281af481f45
ms.date: 06/08/2017
---


# OfficeDataSourceObject Object (Office)

Represents the mail merge data source in a mail merge operation.


## Remarks

To work with the  **OfficeDataSourceObject** object, dimension a variable as an **OfficeDataSourceObject** object. You can then work with the different properties and methods associated with the object. Use the **SetSortOrder** method to specify how to sort the records in a data source.


## Example

 The following example sorts the data source first according to Postal Code in descending order, then on last name and first name in ascending order.


```vb
Sub SetDataSortOrder() 
 Dim appOffice As OfficeDataSourceObject 
 
 Set appOffice = Application.OfficeDataSourceObject 
 appOffice.Open bstrConnect:="DRIVER=SQL Server;SERVER=ServerName;" &amp; _ 
 "UID=user;PWD=;DATABASE=Northwind", bstrTable:="Employees" 
 
 appOffice.SetSortOrder SortField1:="ZipCode", _ 
 SortAscending1:=False, SortField2:="LastName", _ 
 SortField3:="FirstName" 
End Sub
```

Use the  **Column**, **Comparison**, **CompareTo**, and **Conjunction** properties to return or set the data source query criterion. The following example changes an existing filter to remove from the mail merge all records that do not have a Region field equal to "WA".




```vb
Sub SetQueryCriterion() 
 Dim appOffice As Office.OfficeDataSourceObject 
 Dim intItem As Integer 
 
 appOffice.Open bstrConnect:="DRIVER=SQL Server;SERVER=ServerName;" &amp; _ 
 "UID=user;PWD=;DATABASE=Northwind", bstrTable:="Employees" 
 
 With appOffice.Filters 
 For intItem = 1 To .Count 
 With .Item(intItem) 
 If .Column = "Region" Then 
 .Comparison = msoFilterComparisonNotEqual 
 .CompareTo = "WA" 
 If .Conjunction = "Or" Then .Conjunction = "And" 
 End If 
 End With 
 Next intItem 
 End With 
End Sub
```


## Methods



|**Name**|
|:-----|
|[ApplyFilter](Office.OfficeDataSourceObject.ApplyFilter.md)|
|[Move](Office.OfficeDataSourceObject.Move.md)|
|[Open](Office.OfficeDataSourceObject.Open.md)|
|[SetSortOrder](Office.OfficeDataSourceObject.SetSortOrder.md)|

## Properties



|**Name**|
|:-----|
|[Columns](Office.OfficeDataSourceObject.Columns.md)|
|[ConnectString](Office.OfficeDataSourceObject.ConnectString.md)|
|[DataSource](Office.OfficeDataSourceObject.DataSource.md)|
|[Filters](Office.OfficeDataSourceObject.Filters.md)|
|[RowCount](Office.OfficeDataSourceObject.RowCount.md)|
|[Table](Office.OfficeDataSourceObject.Table.md)|

## See also


#### Other resources


[Object Model Reference](./overview/reference-object-library-reference-for-office.md)
