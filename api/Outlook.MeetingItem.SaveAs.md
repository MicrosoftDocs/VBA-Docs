---
title: MeetingItem.SaveAs Method (Outlook)
keywords: vbaol11.chm1435
f1_keywords:
- vbaol11.chm1435
ms.prod: outlook
api_name:
- Outlook.MeetingItem.SaveAs
ms.assetid: cda4cccc-1930-3aa8-d0e1-651de6b0a0b7
ms.date: 06/08/2017
localization_priority: Normal
---


# MeetingItem.SaveAs Method (Outlook)

Saves the Microsoft Outlook item to the specified path and in the format of the specified file type. If the file type is not specified, the MSG format (.msg) is used.


## Syntax

_expression_. `SaveAs`( `_Path_` , `_Type_` )

_expression_ A variable that represents a [MeetingItem](./Outlook.MeetingItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Path_|Required| **String**|The path in which to save the item.|
| _Type_|Optional| **Variant**|The file type to save. Can be one of the following  **OlSaveAsType** constants: **olHTML** , **olMSG** , **olRTF** , **olTemplate** , **olDoc** , ** olTXT** , **olVCal** , **olVCard** , **olICal** , or **olMSGUnicode**.|

## Remarks

Also note that even though  **olDoc** is a valid **OlSaveAsType** constant, messages in HTML format cannot be saved in Document format, and the **olDoc** constant works only if Microsoft Word is set up as the default email editor.


## See also


[MeetingItem Object](Outlook.MeetingItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]