---
title: ProtectedViewWindow.Edit method (Word)
keywords: vbawd10.chm231735397
f1_keywords:
- vbawd10.chm231735397
ms.prod: word
api_name:
- Word.ProtectedViewWindow.Edit
ms.assetid: 8bd4c5cd-8c7a-6bc7-349a-f5ea3d66d921
ms.date: 06/08/2017
localization_priority: Normal
---


# ProtectedViewWindow.Edit method (Word)




## Syntax

 _expression_. `Edit`( `_PasswordTemplate_` , `_WritePasswordDocument_` , `_WritePasswordTemplate_` )

 _expression_ An expression that returns a [ProtectedViewWindow](./Word.ProtectedViewWindow.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _PasswordTemplate_|Optional| **Variant**|The password for opening the template.|
| _WritePasswordDocument_|Optional| **Variant**|The password for saving changes to the document.|
| _WritePasswordTemplate_|Optional| **Variant**|The password for saving changes to the template.|

## Return value

Document


## Remarks

Avoid using hard-coded passwords in your applications. If a password is required in a procedure, request the password from the user, store it in a variable, and then use the variable in your code.


## Example

The following code example opens (for editing) the document associated with the active protected view window.


```vb
Dim pvDoc As Document 
 
Set pvDoc = ActiveProtectedViewWindow.Edit
```


## See also


[ProtectedViewWindow Object](Word.ProtectedViewWindow.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]