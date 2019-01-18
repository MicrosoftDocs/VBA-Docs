---
title: RemoteItem.SaveAs Method (Outlook)
keywords: vbaol11.chm1619
f1_keywords:
- vbaol11.chm1619
ms.prod: outlook
api_name:
- Outlook.RemoteItem.SaveAs
ms.assetid: 1c2c7b68-5239-05f8-4291-d2584fe95194
ms.date: 06/08/2017
localization_priority: Normal
---


# RemoteItem.SaveAs Method (Outlook)

Saves the Microsoft Outlook item to the specified path and in the format of the specified file type. If the file type is not specified, the MSG format (.msg) is used.


## Syntax

_expression_. `SaveAs`( `_Path_` , `_Type_` )

_expression_ A variable that represents a '[RemoteItem](Outlook.RemoteItem.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Path_|Required| **String**|The path in which to save the item.|
| _Type_|Optional| **Variant**|The file type to save. Can be one of the following  **[OlSaveAsType](Outlook.OlSaveAsType.md)** constants: **olHTML** , **olMSG** , **olRTF** , **olTemplate** , **olDoc** , ** olTXT** , **olVCal** , **olVCard** , **olICal** , or **olMSGUnicode**.|

## Remarks

Also note that even though  **olDoc** is a valid **OlSaveAsType** constant, messages in HTML format cannot be saved in Document format, and the **olDoc** constant works only if Microsoft Word is set up as the default email editor.


## See also


[RemoteItem Object](Outlook.RemoteItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]