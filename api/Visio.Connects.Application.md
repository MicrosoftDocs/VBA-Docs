---
title: Connects.Application property (Visio)
keywords: vis_sdr.chm10413090
f1_keywords:
- vis_sdr.chm10413090
ms.prod: visio
api_name:
- Visio.Connects.Application
ms.assetid: add9261d-b2e7-f96f-55c2-8326f8b39813
ms.date: 06/08/2017
localization_priority: Normal
---


# Connects.Application property (Visio)

Returns the instance of Microsoft Visio that is associated with an object. Read-only.


## Syntax

_expression_.**Application**

_expression_ A variable that represents a **[Connects](Visio.Connects.md)** object.


## Return value

**[Application](visio.application.md)**


## Example

The following Microsoft Visual Basic for Applications (VBA) macro gets the  **Application** object associated with the active document and prints its process ID number in the Immediate window.


```vb
 
Public Sub Application_Example() 
 
 Dim vsoApplication As Visio.Application 
 Dim vsoDocument As Visio.Document 
 
 Set vsoDocument = ActiveDocument 
 
 'Get the instance of Visio associated with the Document object. 
 Set vsoApplication = vsoDocument.Application 
 Debug.Print "The process ID of the Application object associated with the active document is: " & vsoApplication.ProcessID 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]