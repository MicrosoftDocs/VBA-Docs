---
title: MailMerge.DataSource property (Publisher)
keywords: vbapb10.chm6225923
f1_keywords:
- vbapb10.chm6225923
ms.prod: publisher
api_name:
- Publisher.MailMerge.DataSource
ms.assetid: 19b32513-fd57-617a-38e2-6230e3e036b9
ms.date: 06/08/2019
localization_priority: Normal
---


# MailMerge.DataSource property (Publisher)

Returns a **[MailMergeDataSource](Publisher.MailMergeDataSource.md)** object that refers to the data source attached to a mail merge or catalog merge main publication.


## Syntax

_expression_.**DataSource**

_expression_ A variable that represents a **[MailMerge](Publisher.MailMerge.md)** object.


## Return value

MailMergeDataSource


## Example

This example displays the path and file name of the data source attached to the active publication.

```vb
Sub DataSourceName() 
 With ActiveDocument.MailMerge.DataSource 
 If .Name <> "" Then _ 
 MsgBox "The path and file name of the " & _ 
 "attached data source is : " & vbCr & .Name 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]