---
title: System.ProfileString property (Word)
keywords: vbawd10.chm154468361
f1_keywords:
- vbawd10.chm154468361
ms.prod: word
api_name:
- Word.System.ProfileString
ms.assetid: c682a0b6-988c-4b81-4314-787fd432afef
ms.date: 06/08/2017
localization_priority: Normal
---


# System.ProfileString property (Word)

Returns or sets a value for an entry in the Windows registry under the following subkey: `HKEY_CURRENT_USER\Software\Microsoft\Office\version\Word`. Read/write  **String**.


## Syntax

 _expression_. `ProfileString`( `_Section_` , `_ Key_` )

 _expression_ An expression that returns a '[System](Word.System.md)' object.


## Example

This example retrieves and displays the startup path stored in the Windows registry.


```vb
MsgBox System.ProfileString("Options", "STARTUP-PATH")
```

This example sets and returns the value for an entry in the Windows registry (the SubkeyName subkey is added below  `HKEY_CURRENT_USER\Software\Microsoft\Office\version\Word`).




```vb
System.ProfileString("SubkeyName", "EntryName") = "Value" 
MsgBox System.ProfileString("SubkeyName", "EntryName")
```


## See also


[System Object](Word.System.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]