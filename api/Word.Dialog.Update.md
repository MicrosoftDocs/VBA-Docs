---
title: Dialog.Update method (Word)
keywords: vbawd10.chm163053870
f1_keywords:
- vbawd10.chm163053870
ms.prod: word
api_name:
- Word.Dialog.Update
ms.assetid: 7adf7403-77eb-85da-8a5a-092d1c8c548f
ms.date: 06/08/2017
localization_priority: Normal
---


# Dialog.Update method (Word)

Updates the values shown in a built-in Microsoft Word dialog box.


## Syntax

_expression_.**Update**

_expression_ Required. A variable that represents a '[Dialog](Word.Dialog.md)' object.


## Example

This example returns a  **Dialog** object that refers to the **Font** dialog box. The font applied to the **Selection** object is changed to Arial, the dialog values are updated, and the **Font** dialog box is displayed.


```vb
Set myDialog = Dialogs(wdDialogFormatFont) 
Selection.Font.Name = "Arial" 
myDialog.Update 
myDialog.Show
```


## See also


[Dialog Object](Word.Dialog.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]