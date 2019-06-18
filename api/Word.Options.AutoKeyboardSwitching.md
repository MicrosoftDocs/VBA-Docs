---
title: Options.AutoKeyboardSwitching property (Word)
keywords: vbawd10.chm162988431
f1_keywords:
- vbawd10.chm162988431
ms.prod: word
api_name:
- Word.Options.AutoKeyboardSwitching
ms.assetid: 22bc427f-20fd-107e-b3c0-c1ec9866a716
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.AutoKeyboardSwitching property (Word)

 **True** if Microsoft Word automatically switches the keyboard language to match what you are typing at any given time. Read/write **Boolean**.


## Syntax

_expression_. `AutoKeyboardSwitching`

_expression_ A variable that represents an **[Options](Word.Options.md)** object.


## Remarks

To use this property, you must have the  **[CheckLanguage](Word.Application.CheckLanguage.md)** property set to **True**.


## Example

This example asks the user to choose whether or not to enable automatic keyboard switching for multilingual documents.


```vb
x = MsgBox("Enable automatic keyboard switching?", vbYesNo) 
If x = vbYes Then 
 Application.CheckLanguage = True 
 Options.AutoKeyboardSwitching = True 
 MsgBox "Automatic keyboard switching enabled!" 
End If
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]