---
title: AddIns.Remove method (PowerPoint)
keywords: vbapp10.chm520005
f1_keywords:
- vbapp10.chm520005
ms.prod: powerpoint
api_name:
- PowerPoint.AddIns.Remove
ms.assetid: 6a7548a4-f7b4-ec80-2cc2-048215913fd4
ms.date: 06/08/2017
localization_priority: Normal
---


# AddIns.Remove method (PowerPoint)

Removes an add-in from the collection of add-ins.


## Syntax

_expression_.**Remove** (_Index_)

_expression_ A variable that represents a [AddIns](PowerPoint.AddIns.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Variant**|The name or index number of the add-in to be removed from the collection.|

## Example

This example removes the add-in named "MyTools" from the list of available add-ins.


```vb
AddIns.Remove "mytools"
```


## See also


[AddIns Object](PowerPoint.AddIns.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]