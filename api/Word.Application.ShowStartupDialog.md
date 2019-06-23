---
title: Application.ShowStartupDialog property (Word)
keywords: vbawd10.chm158335431
f1_keywords:
- vbawd10.chm158335431
ms.prod: word
api_name:
- Word.Application.ShowStartupDialog
ms.assetid: ecee5bb2-271b-f1fc-c25c-a77a59f5df03
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ShowStartupDialog property (Word)

 **True** to display the **Task Pane** when starting Microsoft Word. Read/write **Boolean**.


## Syntax

_expression_. `ShowStartupDialog`

_expression_ A variable that represents an **[Application](Word.Application.md)** object. 


## Remarks

The  **ShowStartupDialog** property is a global option, and the new setting will take effect only after you restart Word. Use the **Visible** property of the **CommandBars** collection show or hide the Task Pane without restarting Word.


## Example

This example turns off the  **Task Pane**, so it won't display upon starting Word. This will not take effect until the next time the user starts Word.


```vb
Sub HideStartUpDlg() 
 Application.ShowStartupDialog = False 
End Sub
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]