---
title: AccessObjectProperties.Add method (Access)
keywords: vbaac10.chm12703
f1_keywords:
- vbaac10.chm12703
ms.prod: access
api_name:
- Access.AccessObjectProperties.Add
ms.assetid: 8f86d5f8-b9af-87d3-fae4-e1a24d7225b6
ms.date: 02/01/2019
localization_priority: Normal
---


# AccessObjectProperties.Add method (Access)

You can use the **Add** method to add a new property as an **[AccessObjectProperty](access.accessobjectproperty.md)** object to the **AccessObjectProperties** collection of an **[AccessObject](Access.AccessObject.md)** object.


## Syntax

_expression_.**Add** (_PropertyName_, _Value_)

_expression_ A variable that represents an **[AccessObjectProperties](Access.AccessObjectProperties.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _PropertyName_|Required|**String**|A string expression that's the name of the new property.|
| _Value_|Required|**Variant**|A **Variant** value corresponding to the option setting. The setting of the value argument depends on the possible settings for a particular option. Can be a constant or a string value.|

## Remarks

You can use the **Remove** method of the **AccessObjectProperties** collection to delete an existing property.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]