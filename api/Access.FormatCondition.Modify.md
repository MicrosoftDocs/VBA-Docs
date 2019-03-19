---
title: FormatCondition.Modify method (Access)
keywords: vbaac10.chm10062
f1_keywords:
- vbaac10.chm10062
ms.prod: access
api_name:
- Access.FormatCondition.Modify
ms.assetid: 213a50f2-30ae-bcdc-d690-2d45bbe6f6e7
ms.date: 03/20/2019
localization_priority: Normal
---


# FormatCondition.Modify method (Access)

You can use the **Modify** method to change the format conditions of a **FormatCondition** object in the **[FormatConditions](Access.FormatConditions.md)** collection of a combo box or text box control.


## Syntax

_expression_.**Modify** (_Type_, _Operator_, _Expression1_, _Expression2_)

_expression_ A variable that represents a **[FormatCondition](Access.FormatCondition.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Type_|Required|**[AcFormatConditionType](Access.AcFormatConditionType.md)**|An **AcFormatConditionType** constant that specifies the type of condition to be modified.|
| _Operator_|Optional|**[AcFormatConditionOperator](access.acformatconditionoperator.md)**|An **AcFormatConditionOperator** constant that specifies the type of operator to be used.<br/><br/>**NOTE**: If the type argument is **acExpression**, the operator argument is ignored. If you leave this argument blank, the default constant (**acBetween**) is assumed. |
| _Expression1_|Optional|**Variant**|A value or expression associated with the first part of the conditional format. Can be a constant value or a string value.|
| _Expression2_|Optional|**Variant**|A value or expression associated with the second part of the conditional format when the operator argument is **acBetween** or **acNotBetween** (otherwise, this argument is ignored). Can be a constant value or a string value.|

## Return value

Nothing



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]