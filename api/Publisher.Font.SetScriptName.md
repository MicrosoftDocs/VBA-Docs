---
title: Font.SetScriptName method (Publisher)
keywords: vbapb10.chm5374001
f1_keywords:
- vbapb10.chm5374001
ms.prod: publisher
api_name:
- Publisher.Font.SetScriptName
ms.assetid: f1f2c01e-098c-1afd-0e64-1d563c1ca626
ms.date: 06/08/2017
localization_priority: Normal
---


# Font.SetScriptName method (Publisher)

Sets the name of the font script to use in a text range.


## Syntax

 _expression_. **SetScriptName**(**_Script_**,  **_FontName_**)

 _expression_ A variable that represents a  **Font** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|Script|Required| **PbFontScriptType**|The script name.|
|FontName|Required| **String**|The font name.|

## Remarks

The Script parameter can be one of the  **[PbFontScriptType](Publisher.PbFontScriptType.md)** constants declared in the Microsoft Publisher type library.


## Example

This example verifies that the default font script in use for the specified text range is Tahoma and, if not, sets it as the default font script.


```vb
Sub GetScript() 
 With ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.Font 
 If .GetScriptName(Script:=pbFontScriptDefault) <> "Tahoma" Then 
 .SetScriptName Script:=pbFontScriptDefault, _ 
 FontName:="Tahoma" 
 End If 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]