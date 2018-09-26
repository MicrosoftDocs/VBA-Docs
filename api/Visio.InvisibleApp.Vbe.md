---
title: InvisibleApp.Vbe Property (Visio)
keywords: vis_sdr.chm17514630
f1_keywords:
- vis_sdr.chm17514630
ms.prod: visio
api_name:
- Visio.InvisibleApp.Vbe
ms.assetid: e8b6a3dc-f3f0-a680-4198-f6b7a82565ae
ms.date: 06/08/2017
---


# InvisibleApp.Vbe Property (Visio)

Gets the root object of the object model exposed by Microsoft Visual Basic for Applications (VBA). Use this property to access and manipulate the VBA projects associated with currently open Microsoft Visio documents. Read-only.


## Syntax

 _expression_. `Vbe`

 _expression_ A variable that represents an [InvisibleApp](./Visio.InvisibleApp.md) object.


### Return value

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


