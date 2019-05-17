---
title: TextFrame.TextRange property (Word)
keywords: vbawd10.chm162661353
f1_keywords:
- vbawd10.chm162661353
ms.prod: word
api_name:
- Word.TextFrame.TextRange
ms.assetid: fd715d4e-6995-2b28-d842-2897d7c1097f
ms.date: 06/08/2017
localization_priority: Normal
---


# TextFrame.TextRange property (Word)

Returns a  **[Range](Word.Range.md)** object that represents the text in the specified text frame.


## Syntax

_expression_.**TextRange**

 _expression_ An expression that returns a **[TextFrame](Word.TextFrame.md)** object.


## Example

This example adds a text box to the active document and then adds text to the text box.


```vb
Set myTBox = ActiveDocument.Shapes _ 
 .AddTextBox(Orientation:=msoTextOrientationHorizontal, _ 
 Left:=100, Top:=100, Width:=300, Height:=200) 
myTBox.TextFrame.TextRange = "Test Box"
```

This example adds text to TextBox 1 in the active document.




```vb
ActiveDocument.Shapes("TextBox 1").TextFrame.TextRange _ 
 .InsertAfter("New Text")
```

This example returns the text from TextBox 1 in the active document and displays it in a message box.




```vb
MsgBox ActiveDocument.Shapes("TextBox 1").TextFrame.TextRange.Text
```


## See also


[TextFrame Object](Word.TextFrame.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]