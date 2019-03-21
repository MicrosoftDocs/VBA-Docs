---
title: ImportExportSpecification object (Access)
keywords: vbaac10.chm13327
f1_keywords:
- vbaac10.chm13327
ms.prod: access
api_name:
- Access.ImportExportSpecification
ms.assetid: a274faba-6da3-35c5-52fc-3341e8def24a
ms.date: 03/21/2019
localization_priority: Normal
---


# ImportExportSpecification object (Access)

Represents a saved import or export operation.


## Remarks

An **ImportExportSpecification** object contains all the information that Microsoft Access needs to repeat an import or export operation without your having to provide any input. 

For example, an import specification that imports data from a Microsoft Office Excel 2007 workbook stores the name of the source Excel file, the name of the destination database, and other details, such as whether you appended to or created a new table, primary key information, field names, and so on.

Use the **[Add](Access.ImportExportSpecifications.Add.md)** method of the **ImportExportSpecifications** collection to create a new **ImportExportSpecification** object.

Use the **Execute** method to run a saved import or export operation.


## Methods

- [Delete](Access.ImportExportSpecification.Delete.md)
- [Execute](Access.ImportExportSpecification.Execute.md)

## Properties

- [Application](Access.ImportExportSpecification.Application.md)
- [Description](Access.ImportExportSpecification.Description.md)
- [Name](Access.ImportExportSpecification.Name.md)
- [Parent](Access.ImportExportSpecification.Parent.md)
- [XML](Access.ImportExportSpecification.XML.md)

## See also

- [Access Object Model Reference](overview/Access/object-model.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
