---
title: Document.ChangeDocument method (Publisher)
keywords: vbapb10.chm196756
f1_keywords:
- vbapb10.chm196756
ms.prod: publisher
api_name:
- Publisher.Document.ChangeDocument
ms.assetid: c6defa92-99fb-973b-6bb2-e3c2a1b0a4f3
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.ChangeDocument method (Publisher)

Changes the current publication to one that uses the wizard, and optionally the design, that you specify.


## Syntax

 _expression_.**ChangeDocument**(**_Wizard_**,  **_Design_**)

 _expression_ A variable that represents a  **Document** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|Wizard|Required| **PbWizard**|The type of wizard. See Remarks for possible values.|
|Design|Optional| **Long**|The design type.|

## Remarks

Possible values for the Wizard parameter are declared in the  **[PbWizard](Publisher.PbWizard.md)** enumeration in the Publisher type library.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **ChangeDocument** method to change the wizard used by the current publication to a brochure.


```vb
Public Sub ChangeDocument_Example() 
 
 ThisDocument.ChangeDocument pbWizardBrochures 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]