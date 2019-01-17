---
title: OfficeDataSourceObject.DataSource property (Office)
keywords: vbaof11.chm232003
f1_keywords:
- vbaof11.chm232003
ms.prod: office
api_name:
- Office.OfficeDataSourceObject.DataSource
ms.assetid: 280e4d63-14d1-8b3c-b735-0190dba12ba1
ms.date: 06/08/2017
localization_priority: Normal
---


# OfficeDataSourceObject.DataSource property (Office)

Gets or sets a  **String** that represents the name of the attached data source. Read/write.


## Syntax

_expression_. `DataSource`

_expression_ A variable that represents an [OfficeDataSourceObject](Office.OfficeDataSourceObject.md) object.


## Example

The following example sets the name of the data source if the name is blank.


```vb
Sub SetAndReturnDataSourceName() 
 Dim appOffice As OfficeDataSourceObject 
 
 Set appOffice = Application.OfficeDataSourceObject 
 With appOffice 
 .Open bstrConnect:="DRIVER=SQL Server;SERVER=ServerName;" &amp; _ 
 "UID=user;PWD=;DATABASE=Northwind", bstrTable:="Employees" 
 
 If .DataSource = "" Then 
 .DataSource = "Northwind" 
 MsgBox .DataSource 
 End If 
 End With 
End Sub
```


## See also


[OfficeDataSourceObject Object](Office.OfficeDataSourceObject.md)



[OfficeDataSourceObject Object Members](./overview/Library-Reference/officedatasourceobject-members-office.md)

