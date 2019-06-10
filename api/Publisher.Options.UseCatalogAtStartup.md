---
title: Options.UseCatalogAtStartup property (Publisher)
keywords: vbapb10.chm1048612
f1_keywords:
- vbapb10.chm1048612
ms.prod: publisher
api_name:
- Publisher.Options.UseCatalogAtStartup
ms.assetid: 7b0cfce9-92f1-5491-c550-421d1c848e0f
ms.date: 06/11/2019
localization_priority: Normal
---


# Options.UseCatalogAtStartup property (Publisher)

**True** for Microsoft Publisher to show the catalog when starting. Read/write **Boolean**.


## Syntax

_expression_.**UseCatalogAtStartup**

_expression_ A variable that represents an **[Options](Publisher.Options.md)** object.


## Return value

Boolean


## Example

This example sets global options for Publisher, including not displaying the catalog upon startup.

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