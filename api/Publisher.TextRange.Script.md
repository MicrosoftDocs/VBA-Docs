---
title: TextRange.Script property (Publisher)
keywords: vbapb10.chm5308484
f1_keywords:
- vbapb10.chm5308484
ms.prod: publisher
api_name:
- Publisher.TextRange.Script
ms.assetid: 54e5a19f-9cb0-0fbc-5ebe-cd4db6c0de8e
ms.date: 06/15/2019
localization_priority: Normal
---


# TextRange.Script property (Publisher)

Returns a **[PbFontScriptType](Publisher.PbFontScriptType.md)** constant that represents the font script for a text range. Read-only.


## Syntax

_expression_.**Script**

_expression_ A variable that represents a **[TextRange](Publisher.TextRange.md)** object.


## Return value

PbFontScriptType


## Remarks

The **Script** property value can be one of the **PbFontScriptType** constants declared in the Microsoft Publisher type library.


## Example

This example displays a message if the font script used in the specified text range is ASCII Latin. This example assumes that there is at least one shape on the first page of the active publication.

```vb
Sub DisplayScriptType() 
 If ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange _ 
 .Script = pbFontScriptAsciiLatin Then 
 MsgBox "The font script you are using is ASCII Latin." 
 End If 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]