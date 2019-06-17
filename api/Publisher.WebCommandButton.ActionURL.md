---
title: WebCommandButton.ActionURL property (Publisher)
keywords: vbapb10.chm3932163
f1_keywords:
- vbapb10.chm3932163
ms.prod: publisher
api_name:
- Publisher.WebCommandButton.ActionURL
ms.assetid: ede9b18f-1be1-9572-9b78-7dbe0817cfe7
ms.date: 06/18/2019
localization_priority: Normal
---


# WebCommandButton.ActionURL property (Publisher)

Returns or sets a **String** that represents the URL of the server-side script to execute in response to a **Submit** button click. Read/write.


## Syntax

_expression_.**ActionURL**

_expression_ A variable that represents a **[WebCommandButton](Publisher.WebCommandButton.md)** object.


## Return value

String


## Remarks

The default value for the **ActionURL** property is `https://example.microsoft.com/~user/ispscript.cgi`. This property is ignored for **Reset** command buttons.


## Example

This example creates a web form **Submit** command button and sets the script path and file name to run when a user chooses the button.

```vb
Sub CreateActionWebButton() 
 With ActiveDocument.Pages(1).Shapes.AddWebControl _ 
 (Type:=pbWebControlCommandButton, Left:=150, _ 
 Top:=150, Width:=75, Height:=36).WebCommandButton 
 .ButtonText = "Submit" 
 .ButtonType = pbCommandButtonSubmit 
 .ActionURL = "https://www.tailspintoys.com/" & _ 
 "scripts/ispscript.cgi" 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]