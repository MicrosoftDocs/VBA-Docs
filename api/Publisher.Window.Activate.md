---
title: Window.Activate method (Publisher)
keywords: vbapb10.chm262162
f1_keywords:
- vbapb10.chm262162
ms.prod: publisher
api_name:
- Publisher.Window.Activate
ms.assetid: 9bd17970-d038-33de-18ad-139bd9fdb8e8
ms.date: 06/18/2019
localization_priority: Normal
---


# Window.Activate method (Publisher)

Activates a window or OLE object.


## Syntax

_expression_.**Activate**

_expression_ A variable that represents a **[Window](Publisher.Window.md)** object.


## Return value

Nothing


## Remarks

Because Publisher runs in a single window, using the **Activate** method with a **Window** object makes Publisher the active application.


## Example

The following example makes Publisher the active application.

```vb
Application.ActiveWindow.Activate
```

<br/>

The following example adds an Excel spreadsheet to the first page of the active publication and activates the spreadsheet for editing.

```vb
Dim shpSheet As Shape 
 
Set shpSheet = ActiveDocument.Pages(1).Shapes.AddOLEObject _ 
 (Left:=72, Top:=72, ClassName:="Excel.Sheet") 
 
shpSheet.OLEFormat.Activate
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]