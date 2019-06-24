---
title: Application.ActivePage property (Visio)
keywords: vis_sdr.chm10013030
f1_keywords:
- vis_sdr.chm10013030
ms.prod: visio
api_name:
- Visio.Application.ActivePage
ms.assetid: 1d0496aa-a6f5-0886-fb8f-8071f95fa333
ms.date: 06/24/2019
localization_priority: Normal
---


# Application.ActivePage property (Visio)

Returns the active **[Page](visio.page.md)** object. Read-only.


## Syntax

_expression_.**ActivePage**

_expression_ A variable that represents an **[Application](Visio.Application.md)** object.


## Return value

Page

## Remarks

The **ActivePage** property returns a **Page** object only when the active window displays a drawing page; otherwise, it returns **Nothing**. To verify that a page is active, use the **Is** operator to compare the **ActivePage** property with **Nothing**.

It is possible to get the active window without qualification from the Microsoft Visio global object, which is automatically available to VBA code that is part of the VBA project of a Visio document. For example, you can use this code: 

```vb
Set vsoPage = ActivePage
```

If your Visual Studio solution includes the [Microsoft.Office.Interop.Visio](https://docs.microsoft.com/visualstudio/vsto/office-primary-interop-assemblies?view=vs-2019) reference, this property maps to the following types:

- **Microsoft.Office.Interop.Visio.IVApplication.ActivePage**
    

## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to get the active page without qualification from the Visio global object, which is automatically available to VBA code that is part of the VBA project of a Visio document.

```vb
 
Public Sub ActivePage_Example() 
  
    Dim vsoPage As Page  
 
    'Find out if a page exists, and if it does, get the page. 
    If Not(ActivePage Is Nothing)  Then 
        Set vsoPage = ActivePage 
        Debug.Print vsoPage.Name 
    Else 
        Debug.Print "No active page." 
    End If   
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]