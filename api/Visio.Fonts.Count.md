---
title: Fonts.Count property (Visio)
keywords: vis_sdr.chm12113330
f1_keywords:
- vis_sdr.chm12113330
ms.prod: visio
api_name:
- Visio.Fonts.Count
ms.assetid: aa428ddf-2a3c-d0e9-231e-ce49e598daf7
ms.date: 06/08/2017
localization_priority: Normal
---


# Fonts.Count property (Visio)

Returns the number of objects in a collection. Read-only.


## Syntax

_expression_.**Count**

_expression_ A variable that represents a **[Fonts](Visio.Fonts.md)** object.


## Return value

Integer


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Count** property to iterate through a **Documents** collection. It displays the names of all the open Microsoft Visio documents in the Immediate window.


```vb
 
Public Sub Count_Example() 
 
 Dim intCounter As Integer 
 Dim vsoDocument As Visio.Document 
 
 For intCounter = 1 To Documents.Count 
 'Get the next open document. 
 Set vsoDocument = Documents.Item(intCounter) 
 
 'Print its name in the Immediate window. 
 Debug.Print vsoDocument.Name 
 Next intCounter 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]