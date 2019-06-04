---
title: Application.WindowPageChange event (Publisher)
keywords: vbapb10.chm268435460
f1_keywords:
- vbapb10.chm268435460
ms.prod: publisher
api_name:
- Publisher.Application.WindowPageChange
ms.assetid: bb636f6e-da4b-7271-9f59-2b7000270c16
ms.date: 06/05/2019
localization_priority: Normal
---


# Application.WindowPageChange event (Publisher)

Occurs when switching the view from one page to another page in a publication.


## Syntax

_expression_.**WindowPageChange** (_Vw_)

_expression_ A variable that represents an **[Application](Publisher.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Vw_|Required| **View**|The new view that includes the page to which the view has been switched.|

## Example

This example changes the view to display the whole page when switching to a new page in a publication. For this example to work, you must place the **WithEvents** declaration in the General Declarations section of a class module and run the InitializeEvents routine.

```vb
Private WithEvents PubApp As Publisher.Application 
 
Sub InitializeEvents() 
 Set PubApp = Publisher.Application 
End Sub 
 
Private Sub PubApp_WindowPageChange(ByVal Vw As View) 
 Vw.Zoom = pbZoomWholePage 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]