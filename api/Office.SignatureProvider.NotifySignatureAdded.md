---
title: SignatureProvider.NotifySignatureAdded method (Office)
keywords: vbaof11.chm287005
f1_keywords:
- vbaof11.chm287005
ms.prod: office
api_name:
- Office.SignatureProvider.NotifySignatureAdded
ms.assetid: 07eb9589-ff67-e54f-9a83-966738c3df58
ms.date: 01/24/2019
localization_priority: Normal
---


# SignatureProvider.NotifySignatureAdded method (Office)

Used to display a dialog box informing the user that the signing process has completed and providing additional functionality for the add-in.


## Syntax

_expression_.**NotifySignatureAdded**(_ParentWindow_, _psigsetup_, _psiginfo_)

_expression_ An expression that returns a **[SignatureProvider](Office.SignatureProvider.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ParentWindow_|Required|**IOleWindow**|Allows the host application to obtain the handle to the window containing the displayed dialog box.|
| _psigsetup_|Required|**SignatureSetup**|Contains initial settings of the signature provider.|
| _psiginfo_|Required|**SignatureInfo**|Contains information about the signature provider add-in.|

## Remarks

This method is called when the signing process has completed. Allows a signature provider add-in the ability to add additional functionality to the add-in. For example, if you wanted to offer an archive service where a user could upload their signed document, you could use this method to initiate that process.


## Example

The following example, written in C#, shows the implementation of the **NotifySignatureAdded** method in a custom signature provider project.


```cs
 public void NotifySignatureAdded(object parentWindow, SignatureSetup sigsetup, SignatureInfo siginfo) 
 { 
 using (Win32WindowFromOleWindow window = new Win32WindowFromOleWindow(parentWindow)) 
 { 
 MessageBox.Show(window, "Signature has been applied", "Signing Ceremony", MessageBoxButtons.OK); 
 } 
 } 

```

> [!NOTE] 
> Signature providers are implemented exclusively in custom COM add-ins and cannot be implemented in Microsoft Visual Basic for Applications (VBA). 


## See also

- [SignatureProvider object members](overview/Library-Reference/signatureprovider-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]