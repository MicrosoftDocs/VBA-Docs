---
title: Shapes.Application property (Visio)
keywords: vis_sdr.chm11313090
f1_keywords:
- vis_sdr.chm11313090
ms.prod: visio
api_name:
- Visio.Shapes.Application
ms.assetid: dfe74ae8-a6e6-d221-0538-ff549e91d2fe
ms.date: 06/08/2017
localization_priority: Normal
---


# Shapes.Application property (Visio)

Returns the instance of Microsoft Visio that is associated with an object. Read-only.


## Syntax

_expression_.**Application**

_expression_ A variable that represents a **[Shapes](Visio.Shapes.md)** object.


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