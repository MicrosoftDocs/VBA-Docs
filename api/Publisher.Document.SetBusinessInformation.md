---
title: Document.SetBusinessInformation method (Publisher)
keywords: vbapb10.chm196757
f1_keywords:
- vbapb10.chm196757
ms.prod: publisher
api_name:
- Publisher.Document.SetBusinessInformation
ms.assetid: 8549f75f-2fb6-6ac6-ecaf-54a0a9b22dc7
ms.date: 06/06/2019
localization_priority: Normal
---


# Document.SetBusinessInformation method (Publisher)

Applies the specified business information set, which consists of a logo image and business contact information (such as the company name and address), to the current publication.


## Syntax

_expression_.**SetBusinessInformation** (_Name_)

_expression_ A variable that represents a **[Document](Publisher.Document.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Name_|Required| **String**|The name of the business information set to be applied.|

## Remarks

Calling the **SetBusinessInformation** method corresponds to selecting a business information set (in the **Select a Business Information set** list), and then choosing the **Update Publication** button in the **Business Information** dialog box (**Edit** menu) in the Microsoft Publisher user interface (UI). 

You must create and edit business information sets in that dialog box before you can use the **SetBusinessInformation** method to apply them programmatically.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the **SetBusinessInformation** method to apply a specific business information set to the current publication. Before you run this code, substitute for `BISetName` the name of a business information set that you have previously created in the Publisher UI.

```vb
Public Sub SetBusinessInformation_Example() 
 
 ThisDocument.SetBusinessInformation "BISetName" 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]