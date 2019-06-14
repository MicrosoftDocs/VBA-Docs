---
title: TextStyle.Font property (Publisher)
keywords: vbapb10.chm5963780
f1_keywords:
- vbapb10.chm5963780
ms.prod: publisher
api_name:
- Publisher.TextStyle.Font
ms.assetid: 80d7177a-fef9-c3fd-f559-94644a2ba0f7
ms.date: 06/15/2019
localization_priority: Normal
---


# TextStyle.Font property (Publisher)

Sets or returns a **[Font](Publisher.Font.md)** object that represents character formatting attributes applied to the specified object. Read/write.


## Syntax

_expression_.**Font**

_expression_ A variable that represents a **[TextStyle](Publisher.TextStyle.md)** object.


## Example

This example selects text and formats the font as bold.

```vb
Sub test2() 
 With Selection.TextRange 
 .Start = 50 
 .End = 150 
 .Font.Bold = msoTrue 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]