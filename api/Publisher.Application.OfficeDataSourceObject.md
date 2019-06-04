---
title: Application.OfficeDataSourceObject property (Publisher)
keywords: vbapb10.chm131123
f1_keywords:
- vbapb10.chm131123
ms.prod: publisher
api_name:
- Publisher.Application.OfficeDataSourceObject
ms.assetid: d7262328-d5b6-6f55-d8c1-e6c072e29e3f
ms.date: 06/05/2019
localization_priority: Normal
---


# Application.OfficeDataSourceObject property (Publisher)

Returns an **[OfficeDataSourceObject](office.officedatasourceobject.md)** object representing the data source in a mail merge or catalog merge operation. Read-only.


## Syntax

_expression_.**OfficeDataSourceObject**

_expression_ A variable that represents an **[Application](Publisher.Application.md)** object.


## Return value

OfficeDataSourceObject


## Example

The following example displays information about the current mail merge data source.

```vb
Dim odsoTemp As Office.OfficeDataSourceObject 
 
Set odsoTemp = Application.OfficeDataSourceObject 
 
With odsoTemp 
 Debug.Print "Connection string: " & .ConnectString 
 Debug.Print "Data source: " & .DataSource 
 Debug.Print "Table: " & .Table 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]