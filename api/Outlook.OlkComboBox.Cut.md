---
title: OlkComboBox.Cut method (Outlook)
keywords: vbaol11.chm1000226
f1_keywords:
- vbaol11.chm1000226
ms.prod: outlook
api_name:
- Outlook.OlkComboBox.Cut
ms.assetid: 4a0a5362-6b85-65e6-797d-9c34652c0980
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkComboBox.Cut method (Outlook)

Removes the contents of the control and copies the contents to the clipboard.


## Syntax

_expression_.**Cut**

_expression_ A variable that represents an [OlkComboBox](Outlook.OlkComboBox.md) object.


## Remarks

The data is copied to the clipboard in unformatted text format, replacing the existing contents of the clipboard.

If the control style is **olComboBoxStyleListBox**, then this method will not cut anything.


## See also


[OlkComboBox Object](Outlook.OlkComboBox.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]