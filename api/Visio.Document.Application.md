---
title: Document.Application property (Visio)
keywords: vis_sdr.chm10513090
f1_keywords:
- vis_sdr.chm10513090
ms.prod: visio
api_name:
- Visio.Document.Application
ms.assetid: 8643d912-21b2-18b4-e0fe-cc6e9db6ae58
ms.date: 06/08/2017
ms.localizationpriority: medium
---


# Document.Application property (Visio)

Returns the instance of Microsoft Visio that is associated with an object. Read-only.


## Syntax

_expression_.**Application**

_expression_ A variable that represents a **[Document](Visio.Document.md)** object.


## Return value

**[Application](visio.application.md)**


## Remarks

If your Visual Studio solution includes the [Microsoft.Office.Interop.Visio](/visualstudio/vsto/office-primary-interop-assemblies?view=vs-2019&preserve-view=true) reference, this property maps to the following types:


- **Microsoft.Office.Interop.Visio.IVDocument.Application**
    

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