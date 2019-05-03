---
title: Application.DirtyObject method (Access)
keywords: vbaac10.chm14676
f1_keywords:
- vbaac10.chm14676
ms.prod: access
api_name:
- Access.Application.DirtyObject
ms.assetid: caf82388-d822-967f-c5f9-0042955ea8d8
ms.date: 02/05/2019
localization_priority: Normal
---


# Application.DirtyObject method (Access)

Marks a form or report as dirty.


## Syntax

_expression_.**DirtyObject** (_ObjectType_, _ObjectName_)

_expression_ A variable that represents an **[Application](Access.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ObjectType_|Required|**[AcObjectType](access.acobjecttype.md)**|Specifies the type of object to mark as dirty. This argument should be set to **acForm** or **acReport**.|
| _ObjectName_|Required|**String**|Specifies the name of the object to mark as dirty.|

## Remarks

The **DirtyObject** method is useful when you are modifying a form or report in a web database programmatically. When you do this, Microsoft Access does not automatically detect that the form or report has changed, which may cause you to lose the changes when you save and close the object. If you use the **DirtyObject** method to specify that the form or report has been changed, the changes will be saved when you save the form or report.

A run-time error occurs if the form or report specified by the _ObjectName_ argument is not open.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]