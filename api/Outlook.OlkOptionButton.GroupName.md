---
title: OlkOptionButton.GroupName property (Outlook)
keywords: vbaol11.chm1000172
f1_keywords:
- vbaol11.chm1000172
ms.prod: outlook
api_name:
- Outlook.OlkOptionButton.GroupName
ms.assetid: 10d091d7-4dae-fa13-abca-424ae27cafa6
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkOptionButton.GroupName property (Outlook)

Returns or sets a **String** that identifies a group of mutually exclusive option button controls. Read/write.


## Syntax

_expression_. `GroupName`

_expression_ A variable that represents an [OlkOptionButton](Outlook.OlkOptionButton.md) object.


## Remarks

Only one member in the group can be selected at a time. Selecting a new member of the group automatically removes the previous selection of any other group member. 

The default value is an empty string.


## See also


[OlkOptionButton Object](Outlook.OlkOptionButton.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]