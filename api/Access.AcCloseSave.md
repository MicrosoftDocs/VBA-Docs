---
title: AcCloseSave enumeration (Access)
keywords: vbaac10.chm10008
f1_keywords:
- vbaac10.chm10008
ms.prod: access
api_name:
- Access.AcCloseSave
ms.assetid: 52cb93d5-8430-7f16-533e-37e981de3829
ms.date: 06/08/2019
localization_priority: Normal
---


# AcCloseSave enumeration (Access)

Used by the **Close** method to specify whether or not to save an object upon closing.

<br/>

|Name|Value|Description|
|:-----|:-----|:-----|
|**acSaveNo**|2|The specified object is not saved.|
|**acSavePrompt**|0|The user is asked whether or not they want to save the object.<br/><br/>**NOTE**: This value is ignored if you are closing a Visual Basic module. The module will be closed, but changes to the module will not be saved.|
|**acSaveYes**|1|The specified object is saved.|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]