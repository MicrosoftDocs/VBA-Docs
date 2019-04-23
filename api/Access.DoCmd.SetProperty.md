---
title: DoCmd.SetProperty method (Access)
keywords: vbaac10.chm5775
f1_keywords:
- vbaac10.chm5775
ms.prod: access
api_name:
- Access.DoCmd.SetProperty
ms.assetid: 32347eb6-115d-36c5-4c18-eab7e7422b78
ms.date: 03/07/2019
localization_priority: Normal
---


# DoCmd.SetProperty method (Access)

The **SetProperty** method carries out the SetProperty action in Visual Basic.


## Syntax

_expression_.**SetProperty** (_ControlName_, _Property_, _Value_)

_expression_ A variable that represents a **[DoCmd](Access.DoCmd.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ControlName_|Required|**Variant**|The name of the field or control for which you want to set the property value. Leave this argument blank to set the property for the current form or report.|
| _Property_|Optional|**Variant**|An **[AcProperty](Access.AcProperty.md)** constant that specifies the property that you want to set.|
| _Value_|Optional|**Variant**|The value to which the property is to be set. For properties whose values are either Yes or No, use 1 for Yes and 0 for No.|



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
