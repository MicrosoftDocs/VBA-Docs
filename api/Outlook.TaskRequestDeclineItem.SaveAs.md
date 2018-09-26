---
title: TaskRequestDeclineItem.SaveAs Method (Outlook)
keywords: vbaol11.chm1856
f1_keywords:
- vbaol11.chm1856
ms.prod: outlook
api_name:
- Outlook.TaskRequestDeclineItem.SaveAs
ms.assetid: 5accd0ea-e9e1-88ec-c614-2b34b7594dd4
ms.date: 06/08/2017
---


# TaskRequestDeclineItem.SaveAs Method (Outlook)

Saves the Microsoft Outlook item to the specified path and in the format of the specified file type. If the file type is not specified, the MSG format (.msg) is used.


## Syntax

 _expression_. `SaveAs`( `_Path_` , `_Type_` )

 _expression_ A variable that represents a [TaskRequestDeclineItem](./Outlook.TaskRequestDeclineItem.md) object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Path_|Required| **String**|The path in which to save the item.|
| _Type_|Optional| **Variant**|The file type to save. Can be one of the following  **OlSaveAsType** constants: **olHTML** , **olMSG** , **olRTF** , **olTemplate** , **olDoc** , ** olTXT** , **olVCal** , **olVCard** , **olICal** , or **olMSGUnicode** .|

## Remarks

Also note that even though  **olDoc** is a valid **OlSaveAsType** constant, messages in HTML format cannot be saved in Document format, and the **olDoc** constant works only if Microsoft Word is set up as the default email editor.


## See also


[TaskRequestDeclineItem Object](Outlook.TaskRequestDeclineItem.md)

