---
title: CodeProject.Properties property (Access)
keywords: vbaac10.chm12721
f1_keywords:
- vbaac10.chm12721
ms.prod: access
api_name:
- Access.CodeProject.Properties
ms.assetid: 47617f8c-6c87-ec70-5661-51204ef44cdf
ms.date: 02/23/2019
localization_priority: Normal
---


# CodeProject.Properties property (Access)

Returns a reference to a **[CodeProject](Access.CodeProject.md)** object's **[AccessObjectProperties](Access.AccessObjectProperties.md)** collection. Read-only.


## Syntax

_expression_.**Properties**

_expression_ A variable that represents a **[CodeProject](Access.CodeProject.md)** object.


## Remarks

The **AccessObjectProperties** collection object is the collection of all the properties related to a **CodeProject** object. You can refer to individual members of the collection by using the member object's index or a string expression that is the name of the member object. 

The first member object in the collection has an index value of 0, and the total number of member objects in the collection is the value of the **AccessObjectProperties** collection's **Count** property minus 1.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]