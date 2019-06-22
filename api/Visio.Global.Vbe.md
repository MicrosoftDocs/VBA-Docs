---
title: Global.VBE property (Visio)
keywords: vis_sdr.chm12814630
f1_keywords:
- vis_sdr.chm12814630
ms.prod: visio
api_name:
- Visio.Global.VBE
ms.assetid: e9840bf6-b6a5-39e5-36c8-517cd772d54d
ms.date: 06/08/2017
localization_priority: Normal
---


# Global.VBE property (Visio)

Gets the root object of the object model exposed by Microsoft Visual Basic for Applications (VBA). Use this property to access and manipulate the VBA projects associated with currently open Microsoft Visio documents. Read-only.


## Syntax

_expression_.**VBE**

_expression_ A variable that represents a **[Global](Visio.Global.md)** object.


## Return value

Object


## Remarks

To get information about the object returned by the  **Vbe** property, follow these steps:


### To get information about the object returned by the Vbe property


1. In the  **Code** group on the [Developer](../visio/How-to/run-visio-in-developer-mode.md) tab, click **Visual Basic**.
    
2. In the Visual Basic Editor, on the  **Tools** menu, click **References**.
    
3. In the  **References** dialog box, click **Microsoft Visual Basic for Applications Extensibility 5.3**, and then click  **OK**.
    
4. On the  **View** menu, click **Object Browser**.
    
5. In the  **Project/Library** list, select the **VBIDE** type library.
    
6. In the  **Classes** list, examine the class named **VBE**.
    
Beginning with Visio 2002, the  **Vbe** property raises an exception if you are running in a security-enhanced environment and your system administrator has blocked access to the VBA object model. There is no user interface or programmatic way to turn this on—the system administrator must turn on (or off) access by setting a Group Policy. This helps protect against viruses that spread by accessing the Visual Basic projects in commonly used templates and injecting the virus code into them.


## Example

This VBA macro shows how to use the  **Vbe** property to determine how many VBA projects are open in an instance of Visio.

Before running this code, make sure the  **Trust access to the VBA project object model** check box is selected under **Developer Macro Settings** on the **Macro Settings** page of the **Trust Center** dialog box (click the **File** tab, click **Options**, click  **Trust Center**, and then click  **Trust Center Settings**). 




```vb
 
Public Sub Vbe_Example() 
 
     Dim vbideVBE As VBIDE.VBE 
 
     Set vbideVBE = Visio.Application.Vbe 
     Debug.Print vbideVBE.VBProjects.Count 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]