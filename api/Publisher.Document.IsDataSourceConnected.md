---
title: Document.IsDataSourceConnected property (Publisher)
keywords: vbapb10.chm196722
f1_keywords:
- vbapb10.chm196722
ms.prod: publisher
api_name:
- Publisher.Document.IsDataSourceConnected
ms.assetid: b62422ab-12f7-1151-d8d1-1cb32de18160
ms.date: 06/06/2019
localization_priority: Normal
---


# Document.IsDataSourceConnected property (Publisher)

**True** if the specified publication is connected to a data source. Read-only.


## Syntax

_expression_.**IsDataSourceConnected**

_expression_ A variable that represents a **[Document](Publisher.Document.md)** object.


## Remarks

A publication must be connected to a valid data source to perform a mail merge or catalog merge.


## Example

The following example tests whether the publication is connected to a data source, and if it is not, specifies and connects a data source to the publication. 

Before running this example, you must replace `PathToFile` with a valid file path and `TableName` with a valid data source table name.

```vb
Dim strDataSource As String 
Dim strDataSourceTable As String 
 
 'Specify data source and table name 
 
 strDataSource = "PathToFile" 
 strDataSourceTable = "TableName" 
 
 'Connect to a datasource 
 If Not (ThisDocument.IsDataSourceConnected) Then 
 ThisDocument.MailMerge.OpenDataSource strDataSource, , strDataSourceTable 
 
 End If
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]