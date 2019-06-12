---
title: Options.DragAndDropText property (Publisher)
keywords: vbapb10.chm1048584
f1_keywords:
- vbapb10.chm1048584
ms.prod: publisher
api_name:
- Publisher.Options.DragAndDropText
ms.assetid: 55fb68e8-4ddc-6866-00d8-bdd6a1e25ec3
ms.date: 06/11/2019
localization_priority: Normal
---


# Options.DragAndDropText property (Publisher)

**True** to enable dragging of text. Read/write **Boolean**.


## Syntax

_expression_.**DragAndDropText**

_expression_ A variable that represents an **[Options](Publisher.Options.md)** object.


## Return value

Boolean


## Example

This example sets global options for Microsoft Publisher, including enabling dragging to reposition text.

```vb
Sub SetGlobalOptions() 
 With Options 
 .AutoFormatWord = True 
 .AutoKeyboardSwitching = True 
 .AutoSelectWord = True 
 .DragAndDropText = True 
 .UseCatalogAtStartup = False 
 .UseHelpfulMousePointers = False 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]