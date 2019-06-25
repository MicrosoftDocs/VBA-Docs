---
title: Application.COMAddIns property (Visio)
keywords: vis_sdr.chm10050535
f1_keywords:
- vis_sdr.chm10050535
ms.prod: visio
api_name:
- Visio.Application.COMAddIns
ms.assetid: 182ea1e1-f896-f619-1bf0-df4a57b39abf
ms.date: 06/25/2019
localization_priority: Normal
---


# Application.COMAddIns property (Visio)

Returns a reference to the **[COMAddIns](office.comaddins.md)** collection that represents all the Component Object Model (COM) add-ins currently registered in Microsoft Visio. Read-only.


## Syntax

_expression_.**COMAddIns**

_expression_ A variable that represents an **[Application](Visio.Application.md)** object.


## Return value

Object


## Remarks

The COM add-ins that are currently registered are listed in the **COM Add-Ins** dialog box (**File** tab > **Options** > **Add-Ins** > **Go**)

To get information about the object returned by the **COMAddIns** property:

1. In the **Code** group on the **Developer** tab (**File** tab > **Options** > **Advanced** > **General** > **Run in developer mode**), choose **Visual Basic**.
    
2. On the **View** menu, choose **Object Browser**.
    
3. In the **Project/Library** list, choose **Office**.
    
4. Under **Classes**, examine the class named **COMAddIns**.
    

## Example

This macro shows how to use the **COMAddIns** property to list the COM add-ins registered with Visio.

```vb
 
Public Sub COMAddIns_Example()  
 
    Dim vsoCOMAddIns As COMAddIns  
    Dim vsoCOMAddIn As COMAddIn  
 
    'Get the set of COM add-ins.  
    Set vsoCOMAddIns = Application.COMAddIns  
 
    'List each COM add-in in the Immediate window. 
    For Each vsoCOMAddIn In vsoCOMAddIns  
        Debug.Print vsoCOMAddIn.Description  
    Next 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]