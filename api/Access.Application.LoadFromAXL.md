---
title: Application.LoadFromAXL method (Access)
keywords: vbaac10.chm14665
f1_keywords:
- vbaac10.chm14665
ms.prod: access
api_name:
- Access.Application.LoadFromAXL
ms.assetid: 1cce0568-1966-c089-a741-b0934b8676d6
ms.date: 02/05/2019
localization_priority: Normal
---


# Application.LoadFromAXL method (Access)

Imports the object defined in an Application XML (AXL) file into the database. 


## Syntax

_expression_.**LoadFromAXL** (_ObjectType_, _ObjectName_, _FileName_)

_expression_ A variable that represents an **[Application](Access.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ObjectType_|Required|**AcObjectType**|Specifies the type of object to create.|
| _ObjectName_|Required|**String**|Specifies the name of the object.|
| _FileName_|Required|**String**|Specifies the full path and file name of the AXL file to import.|

## Remarks

The **LoadFromAXL** method does not provide a warning when the object specified in the _ObjectName_ argument already exists. If an object of the same name already exists, it will be replaced by the object specified in the _ObjectName_ argument.

For more information about AXL, see [[MS-AXL]: Access Application Transfer Protocol Structure Specification](https://msdn.microsoft.com/library/dd927584.aspx).




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]