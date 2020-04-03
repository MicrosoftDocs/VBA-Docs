---
title: FormDescription.DisplayName property (Outlook)
keywords: vbaol11.chm187
f1_keywords:
- vbaol11.chm187
ms.prod: outlook
api_name:
- Outlook.FormDescription.DisplayName
ms.assetid: 2b621bd4-2d27-e15b-4c1b-c9a84328abc0
ms.date: 06/08/2017
localization_priority: Normal
---


# FormDescription.DisplayName property (Outlook)

Returns or sets a  **String** representing the name of the form, which is displayed in the **Choose Forms** dialog box. Read/write.


## Syntax

_expression_.**DisplayName**

_expression_ A variable that represents a [FormDescription](Outlook.FormDescription.md) object.


## Remarks

If both the  **[FormDescription.Name](Outlook.FormDescription.Name.md)** and **FormDescription.DisplayName** properties are empty, setting one will set the other. If one has been previously set, setting the other will not change the value.


## See also


[FormDescription Object](Outlook.FormDescription.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]