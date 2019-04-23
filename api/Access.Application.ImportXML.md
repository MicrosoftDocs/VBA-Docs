---
title: Application.ImportXML method (Access)
keywords: vbaac10.chm12604
f1_keywords:
- vbaac10.chm12604
ms.prod: access
api_name:
- Access.Application.ImportXML
ms.assetid: c7baa4be-4ef6-c886-3cd6-06576563b77d
ms.date: 02/05/2019
localization_priority: Normal
---


# Application.ImportXML method (Access)

The **ImportXML** method allows developers to import XML data and/or schema information into Microsoft SQL Server 2000 Desktop Engine (MSDE 2000), Microsoft SQL Server 7.0 or later, or the Microsoft Access database engine.


## Syntax

_expression_.**ImportXML** (_DataSource_, _ImportOptions_)

_expression_ A variable that represents an **[Application](Access.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _DataSource_|Required|**String**|The name and path of the XML file to import.|
| _ImportOptions_|Optional|**[AcImportXMLOption](Access.AcImportXMLOption.md)**|An **AcImportXMLOption** constant that specifies the option to use when importing XML files. The default value is **acStructureAndData**.|

## Return value

Nothing


## Example

The following example imports an XML file into a new table named Employees in the current database.

```vb
Application.ImportXML _ 
 DataSource:="employees.xml", _ 
 ImportOptions:=acStructureAndData
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
