---
title: Application.SetHiddenAttribute method (Access)
keywords: vbaac10.chm12571
f1_keywords:
- vbaac10.chm12571
ms.prod: access
api_name:
- Access.Application.SetHiddenAttribute
ms.assetid: b92a1edc-033a-095c-980f-852b8f7e0785
ms.date: 02/05/2019
localization_priority: Normal
---


# Application.SetHiddenAttribute method (Access)

The **SetHiddenAttribute** method sets the hidden attribute of an Access object.


## Syntax

_expression_.**SetHiddenAttribute** (_ObjectType_, _ObjectName_, _fHidden_)

_expression_ A variable that represents an **[Application](Access.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ObjectType_|Required|**[AcObjectType](Access.AcObjectType.md)**|An **AcObjectType** constant that specifies the type of Access object.|
| _ObjectName_|Required|**String**|The name of the Access object.|
| _fHidden_|Required|**Boolean**|**True** sets the hidden attribute, and **False** clears the attribute.|

## Return value

Nothing


## Remarks

Together with the **[GetHiddenAttribute](access.application.gethiddenattribute.md)** method, the **SetHiddenAttribute** method provides a means of changing an object's visibility from Visual Basic code. With these methods, you can set or read the **Hidden** property available in the object's **Properties** dialog box.

To set this option by using the **SetHiddenAttribute** method, specify **True** or **False** for the setting, as in the following example.

```vb
Application.SetHiddenAttribute acTable,"Customers", True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]