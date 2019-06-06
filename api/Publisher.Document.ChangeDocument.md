---
title: Document.ChangeDocument method (Publisher)
keywords: vbapb10.chm196756
f1_keywords:
- vbapb10.chm196756
ms.prod: publisher
api_name:
- Publisher.Document.ChangeDocument
ms.assetid: c6defa92-99fb-973b-6bb2-e3c2a1b0a4f3
ms.date: 06/06/2019
localization_priority: Normal
---


# Document.ChangeDocument method (Publisher)

Changes the current publication to one that uses the wizard, and optionally the design, that you specify.


## Syntax

_expression_.**ChangeDocument** (_Wizard_, _Design_)

_expression_ A variable that represents a **[Document](Publisher.Document.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Wizard_|Required| **[PbWizard](Publisher.PbWizard.md)**|The type of wizard. Can be one of the **PbWizard** constants.|
|_Design_|Optional| **Long**|The design type.| 


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the **ChangeDocument** method to change the wizard used by the current publication to a brochure.

```vb
Public Sub ChangeDocument_Example() 
 
 ThisDocument.ChangeDocument pbWizardBrochures 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]