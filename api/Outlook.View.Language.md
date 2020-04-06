---
title: View.Language property (Outlook)
keywords: vbaol11.chm2489
f1_keywords:
- vbaol11.chm2489
ms.prod: outlook
api_name:
- Outlook.View.Language
ms.assetid: caa2eb1b-26e3-e8da-c0d8-118d9ba654dc
ms.date: 06/08/2017
localization_priority: Normal
---


# View.Language property (Outlook)

Returns or sets a  **String** value that represents the language setting for the object that defines the language used in the menu. Read/write.


## Syntax

_expression_.**Language**

_expression_ A variable that represents a [View](Outlook.View.md) object.


## Remarks

The  **Language** property uses a **String** to represent an ISO language tag. For example, the string "EN-US" represents the ISO code for "United States - English."

If a valid language code is specified, the object will only be available in the  **View** menu for the specified language type. If no value is specified, the object item is available for all language types. The default value for this property is an empty string.


## See also


[View Object](Outlook.View.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]