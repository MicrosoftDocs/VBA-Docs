---
title: Application.GetObjectReference method (Outlook)
keywords: vbaol11.chm734
f1_keywords:
- vbaol11.chm734
ms.prod: outlook
api_name:
- Outlook.Application.GetObjectReference
ms.assetid: 426ade68-155b-9076-b3f8-4108f44688b0
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.GetObjectReference method (Outlook)

Creates a strong or weak object reference for a specified Outlook object.


## Syntax

_expression_. `GetObjectReference`( `_Item_` , `_ReferenceType_` )

 _expression_ An expression that returns an **[Application](Outlook.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Item_|Required| **Object**|The object from which to obtain a strong or weak object reference.|
| _ReferenceType_|Required| **[OlReferenceType](Outlook.OlReferenceType.md)**|The type of object reference.|

## Return value

An **Object** that represents a strong or weak object reference for the specified object.


## Remarks

This method returns a weak or strong object reference for the object specified in  _Item_.


> [!NOTE] 
> Outlook can fail to close successfully if an add-in retains strong object references. Always dereference a strong object reference once it is no longer needed by the add-in.


## See also


[Application Object](Outlook.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]