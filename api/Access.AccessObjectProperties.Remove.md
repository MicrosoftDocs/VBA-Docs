---
title: AccessObjectProperties.Remove method (Access)
keywords: vbaac10.chm12704
f1_keywords:
- vbaac10.chm12704
ms.prod: access
api_name:
- Access.AccessObjectProperties.Remove
ms.assetid: c06fff7c-2e68-1955-f151-27e105e4be6a
ms.date: 02/01/2019
localization_priority: Normal
---


# AccessObjectProperties.Remove method (Access)

You can use the **Remove** method to remove an **[AccessObjectProperty](access.accessobjectproperty.md)** object from the **AccessObjectProperties** collection of an **[AccessObject](Access.AccessObject.md)** object.


## Syntax

_expression_.**Remove** (_Item_)

_expression_ A variable that represents an **[AccessObjectProperties](Access.AccessObjectProperties.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Item_|Required|**Variant**|An expression that specifies the position of a member of the collection referred to by the object argument.<br/><br/>If a numeric expression, the index argument must be a number from 0 to the value of the collection's **Count** property minus 1.<br/><br/>If a string expression, the index argument must be the name of a member of the collection.|




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]