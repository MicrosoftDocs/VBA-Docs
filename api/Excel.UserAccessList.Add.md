---
title: UserAccessList.Add method (Excel)
keywords: vbaxl10.chm726075
f1_keywords:
- vbaxl10.chm726075
ms.prod: excel
api_name:
- Excel.UserAccessList.Add
ms.assetid: dd3b3bc4-8618-b680-7409-c431a12374b0
ms.date: 05/18/2019
localization_priority: Normal
---


# UserAccessList.Add method (Excel)

Adds a user access list.


## Syntax

_expression_.**Add** (_Name_, _AllowEdit_)

_expression_ A variable that represents a **[UserAccessList](Excel.UserAccessList.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the user access list.|
| _AllowEdit_|Required| **Boolean**| **True** allows users on the access list to edit the editable ranges on a protected worksheet.|

## Return value

A **[UserAccess](Excel.UserAccess.md)** object that represents the new user access list.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]