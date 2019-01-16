---
title: OfficeDataSourceObject.Columns property (Office)
keywords: vbaof11.chm232004
f1_keywords:
- vbaof11.chm232004
ms.prod: office
api_name:
- Office.OfficeDataSourceObject.Columns
ms.assetid: 02a3eb37-df7a-923a-6a98-dbb980b413f7
ms.date: 06/08/2017
localization_priority: Normal
---


# OfficeDataSourceObject.Columns property (Office)

Gets an  **ODSOColumns** object that represents the fields in a data source. Read-only.


## Syntax

_expression_. `Columns`

_expression_ A variable that represents an [OfficeDataSourceObject](Office.OfficeDataSourceObject.md) object.


## Example

The following example displays the field names in the data source attached to the active publication.


```vb
Sub ShowFieldNames() 
 Dim appOffice As OfficeDataSourceObject 
 Dim intCount As Integer 
 
 Set appOffice = Application.OfficeDataSourceObject 
 appOffice.Open bstrConnect:="DRIVER=SQL Server;SERVER=ServerName;" &amp; _ 
 "UID=user;PWD=;DATABASE=Northwind", bstrTable:="Employees" 
 
 With appOffice.Columns 
 For intCount = 1 To .Count 
 MsgBox "Field Name: " &amp; .Item(intCount).Name 
 Next 
 End With 
End Sub
```


## See also


[OfficeDataSourceObject Object](Office.OfficeDataSourceObject.md)



[OfficeDataSourceObject Object Members](./overview/Library-Reference/officedatasourceobject-members-office.md)

