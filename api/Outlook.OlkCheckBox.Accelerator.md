---
title: OlkCheckBox.Accelerator property (Outlook)
keywords: vbaol11.chm1000134
f1_keywords:
- vbaol11.chm1000134
ms.prod: outlook
api_name:
- Outlook.OlkCheckBox.Accelerator
ms.assetid: 2604a27f-472b-ccc6-ad37-317ea0008a39
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkCheckBox.Accelerator property (Outlook)

Returns or sets a  **String** value that represents the accelerator or hot key for the control. Read/write.


## Syntax

_expression_. `Accelerator`

_expression_ A variable that represents an [OlkCheckBox](Outlook.OlkCheckBox.md) object.


## Remarks

The default value is an empty string, meaning there is no hot key. If the property is set with a string of more than one character, the value will be truncated to the first character. 

You cannot use digits in an accelerator.


## See also


[OlkCheckBox Object](Outlook.OlkCheckBox.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]