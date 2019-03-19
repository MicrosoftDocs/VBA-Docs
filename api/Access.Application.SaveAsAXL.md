---
title: Application.SaveAsAXL method (Access)
keywords: vbaac10.chm14664
f1_keywords:
- vbaac10.chm14664
ms.prod: access
api_name:
- Access.Application.SaveAsAXL
ms.assetid: a9557499-7e69-b405-8e2f-d9fcb23fb012
ms.date: 02/05/2019
localization_priority: Normal
---


# Application.SaveAsAXL method (Access)

Exports the specified object to an Application XML (AXL) file.


## Syntax

_expression_.**SaveAsAXL** (_ObjectType_, _ObjectName_, _FileName_)

_expression_ A variable that represents an **[Application](Access.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ObjectType_|Required|**[AcObjectType](access.acobjecttype.md)**|Specifies the type of object to export.|
| _ObjectName_|Required|**String**|Specifies the name of the object to export. |
| _FileName_|Required|**String**|Specifies the full path and file name of the AXL file to create.|

## Remarks

The **SaveAsAXL** method does not provide a warning when the file specified in the _FileName_ argument already exists. If this occurs, the file will be overwritten.

The **SaveAsAXL** method generates a run-time error if the current database is not a web database.

For more information about AXL, see [[MS-AXL]: Access Application Transfer Protocol Structure Specification](https://msdn.microsoft.com/library/dd927584.aspx).




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]