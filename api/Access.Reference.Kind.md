---
title: Reference.Kind property (Access)
keywords: vbaac10.chm12637
f1_keywords:
- vbaac10.chm12637
ms.prod: access
api_name:
- Access.Reference.Kind
ms.assetid: 51a941e2-25c5-3699-232c-c6fb90228f65
ms.date: 03/23/2019
localization_priority: Normal
---


# Reference.Kind property (Access)

The **Kind** property indicates the type of reference that a **Reference** object represents. Read-only **vbext_RefKind**.


## Syntax

_expression_.**Kind**

_expression_ A variable that represents a **[Reference](Access.Reference.md)** object.


## Remarks

The **Kind** property returns the following values.

|Value|Description|
|:-----|:-----|
|**vbext_rk_Project**|The **Reference** object represents a reference to a Visual Basic project.|
|**vbext_rk_TypeLib**|The **Reference** object represents a reference to a file that contains a type library.|



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]