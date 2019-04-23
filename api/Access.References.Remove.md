---
title: References.Remove method (Access)
keywords: vbaac10.chm12644
f1_keywords:
- vbaac10.chm12644
ms.prod: access
api_name:
- Access.References.Remove
ms.assetid: ebdc9da2-cc32-6169-994a-1041b1c49031
ms.date: 03/23/2019
localization_priority: Normal
---


# References.Remove method (Access)

The **Remove** method removes a **[Reference](Access.Reference.md)** object from the **References** collection.


## Syntax

_expression_.**Remove** (_Reference_)

_expression_ A variable that represents a **[References](Access.References.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Reference_|Required|**Reference**|The **Reference** object that represents the reference that you wish to remove.|

## Remarks

To determine the name of the **Reference** object that you wish to remove, check the Project/Library box in the Object Browser. The names of all references that are currently set appear there. These names correspond to the value of the **Name** property of a **Reference** object.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]