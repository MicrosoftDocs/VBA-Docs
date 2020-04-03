---
title: UserDefinedProperties.Find method (Outlook)
keywords: vbaol11.chm589
f1_keywords:
- vbaol11.chm589
ms.prod: outlook
api_name:
- Outlook.UserDefinedProperties.Find
ms.assetid: 1f4ddf1f-b36d-e852-17ff-700708893a30
ms.date: 06/08/2017
localization_priority: Normal
---


# UserDefinedProperties.Find method (Outlook)

Locate a  **[UserDefinedProperty](Outlook.UserDefinedProperty.md)** contained in the collection.


## Syntax

_expression_.**Find** (_Name_)

_expression_ A variable that represents a [UserDefinedProperties](Outlook.UserDefinedProperties.md) object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The  **[Name](Outlook.UserDefinedProperty.Name.md)** property value of the **UserDefinedProperty** object to find in the collection.|

## Return value

A  **UserDefinedProperty** object that represents the located object, if successful; otherwise, **Null** (**Nothing** in Visual Basic).


## See also


[UserDefinedProperties Object](Outlook.UserDefinedProperties.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]