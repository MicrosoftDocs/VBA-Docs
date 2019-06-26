---
title: Shape.ChangePicture method (Visio)
ms.prod: visio
ms.assetid: 9193d802-cebd-2bfd-5f8e-400fac36c1a5
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.ChangePicture method (Visio)

Replaces the specified shape's current picture with a new picture.


## Syntax

_expression_.**ChangePicture** (_FileName_, _ChangePictureFlags_)

_expression_ A variable that represents a **[Shape](Visio.Shape.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FileName_|Required|**String**|Specifies the full path of the replacement picture.|
| _ChangePictureFlags_|Optional|INT32|Reserved for future implementation. Has no effect.|

## Return value

**DOUBLE**


## Remarks

The **DOUBLE** returned represents the ratio of the picture's width to its height.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]