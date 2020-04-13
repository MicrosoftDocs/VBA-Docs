---
title: Global.FindKey property (Word)
keywords: vbawd10.chm163119175
f1_keywords:
- vbawd10.chm163119175
ms.prod: word
api_name:
- Word.Global.FindKey
ms.assetid: 79203ae9-dcc9-ffb1-d974-0eb814268d6e
ms.date: 06/08/2017
localization_priority: Normal
---


# Global.FindKey property (Word)

Returns a  **[KeyBinding](Word.KeyBinding.md)** object that represents the specified key combination. Read-only.


## Syntax

_expression_. `FindKey`( `_KeyCode_` , `_KeyCode2_` )

_expression_ A variable that represents a '[Global](Word.Global.md)' object. Optional.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _KeyCode_|Required| **Long**|A key you specify by using one of the **WdKey** constants.|
| _KeyCode2_|Optional| **Variant**|A second key you specify by using one of the **WdKey** constants.|

## Remarks

You can use the **BuildKeyCode** method to create the KeyCode or KeyCode2 argument.


## Example

This example disables the ALT+SHIFT+F12 key combination in the template attached to the active document. To return a  **KeyBinding** object that includes more than two keys, use the **BuildKeyCode** method, as shown in the example.


```vb
CustomizationContext = ActiveDocument.AttachedTemplate 
FindKey(KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyShift, _ 
 wdKeyF12)).Disable
```

This example displays the command assigned to the F1 key.




```vb
CustomizationContext = NormalTemplate 
MsgBox FindKey(KeyCode:=wdKeyF1).Command
```


## See also


[Global Object](Word.Global.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]