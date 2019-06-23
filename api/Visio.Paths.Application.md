---
title: Paths.Application property (Visio)
keywords: vis_sdr.chm15313090
f1_keywords:
- vis_sdr.chm15313090
ms.prod: visio
api_name:
- Visio.Paths.Application
ms.assetid: d5439015-9efb-46ea-49a8-7f08c4d82a14
ms.date: 06/08/2017
localization_priority: Normal
---


# Paths.Application property (Visio)

Returns the instance of Microsoft Visio that is associated with an object. Read-only.


## Syntax

_expression_.**Application**

_expression_ A variable that represents a **[Paths](Visio.Paths.md)** object.


## Return value

**[Application](visio.application.md)**


## Example

The following Microsoft Visual Basic for Applications (VBA) macro gets the **Application** object associated with the active document and prints its process ID number in the Immediate window.

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