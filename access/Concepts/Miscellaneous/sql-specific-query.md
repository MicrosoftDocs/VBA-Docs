---
title: SQL specific query
ROBOTS: INDEX
keywords: vbaac10.chm47377
f1_keywords:
- vbaac10.chm47377
ms.prod: access
ms.assetid: 506c45eb-c48e-94de-60cd-10058860b3a6
ms.date: 06/08/2017
localization_priority: Normal
---


# SQL specific query

**Applies to:** Access 2013 | Access 2016

An SQL specific query is one that can be created only by writing an SQL statement in SQL view. Union, pass-through, and data definition queries are SQL specific queries.

|**SQL specific query type**|**Description**|
|:-----|:-----|
|Union|An SQL specific select query that combines corresponding fields from two or more tables or queries into one field.<br/><br/>For example, a union query of the Customers table and the Suppliers table results in a snapshot that contains all the specified records from both the Customers table and the Suppliers table.|
|Pass-Through|An SQL specific query that sends commands directly to an SQL database server (such as Microsoft SQL Server).<br/><br/>With pass-through queries, you work with the tables on the server instead of linking the tables to your Microsoft Access database.|
|Data Definition|An SQL specific query that can create or delete an index, or create, alter, or delete a table.|

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]