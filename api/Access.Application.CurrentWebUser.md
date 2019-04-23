---
title: Application.CurrentWebUser method (Access)
keywords: vbaac10.chm14599
f1_keywords:
- vbaac10.chm14599
ms.prod: access
api_name:
- Access.Application.CurrentWebUser
ms.assetid: cb8b230d-71c5-c73d-c88e-1a13246492a5
ms.date: 02/05/2019
localization_priority: Normal
---


# Application.CurrentWebUser method (Access)

Gets information about the current user of a web database on Microsoft SharePoint Foundation 2010 and later.


## Syntax

_expression_.**CurrentWebUser** (_DisplayOption_)

_expression_ A variable that represents an **[Application](Access.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _DisplayOption_|Required|**[AcWebUserDisplay](access.acwebuserdisplay.md)**|Specifies the type of information to return about the user.|

## Return value

Variant


## Remarks

The **CurrentWebUser** method returns **Null** if the database has not been published, or information about the current user cannot be retrieved.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]