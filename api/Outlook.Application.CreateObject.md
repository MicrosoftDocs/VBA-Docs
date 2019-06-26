---
title: Application.CreateObject method (Outlook)
keywords: vbaol11.chm716
f1_keywords:
- vbaol11.chm716
ms.prod: outlook
api_name:
- Outlook.Application.CreateObject
ms.assetid: 09b6ff5b-a750-c07d-7499-c1f8a00214fe
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.CreateObject method (Outlook)

Creates an automation object of the specified class.


## Syntax

_expression_. `CreateObject`( `_ObjectName_` )

_expression_ A variable that represents an **[Application](Outlook.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ObjectName_|Required| **String**|The class name of the object to create. For information about valid class names, see [OLE Programmatic Identifiers](../outlook/Concepts/Getting-Started/ole-programmatic-identifiers-outlook.md).|

## Return value

An Object value that represents the new Automation object instance. If the application is already running, **CreateObject** will create a new instance.


## Remarks

This method is provided so that other applications can be automated from Microsoft Visual Basic Scripting Edition (VBScript) 1.0, which did not include a **CreateObject** method. **CreateObject** has been included in VBScript version 2.0 and later. This method should not be used to automate Microsoft Outlook from VBScript.


> [!NOTE] 
> The **CreateObject** methods commonly used in the example code within this Help file (available when you click "Example") are made available by Microsoft Visual Basic or Microsoft Visual Basic for Applications (VBA). These examples do not use the same **CreateObject** method that is implemented as part of the object model in Outlook.


## Example

This VBScript example uses the **[Open](Outlook.MailItem.Open.md)** event of the item to access Windows Internet Explorer and display the webpage.


```vb
Sub Item_Open() 
 
 Set Web = CreateObject("InternetExplorer.Application") 
 
 Web.Visible = True 
 
 Web.Navigate "www.microsoft.com" 
 
End Sub
```

This VBScript example uses the **Click** event of a **CommandButton** control on the item to access Microsoft Word and open a document in the root directory named "Resume.doc".




```vb
Sub CommandButton1_Click() 
 
 Set Word = Application.CreateObject("Word.Application") 
 
 Word.Visible = True 
 
 Word.Documents.Open("C:\Resume.doc") 
 
End Sub
```


## See also


[Application Object](Outlook.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
