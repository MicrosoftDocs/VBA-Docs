---
title: View.RevisionsBalloonShowConnectingLines property (Word)
keywords: vbawd10.chm161808428
f1_keywords:
- vbawd10.chm161808428
ms.prod: word
api_name:
- Word.View.RevisionsBalloonShowConnectingLines
ms.assetid: 78c1cf42-93a7-eec9-84f6-40c6e7de036c
ms.date: 06/08/2017
localization_priority: Normal
---


# View.RevisionsBalloonShowConnectingLines property (Word)

 **True** for Microsoft Word to display connecting lines from the text to the revision and comment balloons. Read/write **Boolean**.


## Syntax

_expression_. `RevisionsBalloonShowConnectingLines`

_expression_ A variable that represents a '[View](Word.View.md)' object.


## Example

This example hides the lines connecting the document text with the corresponding revision or comment balloons.


```vb
Sub ShowConnectingLines() 
 ActiveWindow.View _ 
 .RevisionsBalloonShowConnectingLines = False 
End Sub
```


## See also


[View Object](Word.View.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]