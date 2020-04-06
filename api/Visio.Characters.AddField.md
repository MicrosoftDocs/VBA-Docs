---
title: Characters.AddField method (Visio)
keywords: vis_sdr.chm10216030
f1_keywords:
- vis_sdr.chm10216030
ms.prod: visio
api_name:
- Visio.Characters.AddField
ms.assetid: 1b00cad3-d97a-4bdc-1f8e-cee39d9c836f
ms.date: 06/08/2017
localization_priority: Normal
---


# Characters.AddField method (Visio)

Replaces the text represented by a  **Characters** object with a new field of the category, code, and format you specify.


## Syntax

_expression_.**AddField** (_Category_, _Code_, _Format_)

_expression_ A variable that represents a **[Characters](Visio.Characters.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Category_|Required| **Integer**| **VisFieldCategories**. The category for the new field.|
| _Code_|Required| **Integer**| **VisFieldCodes**. The code for the new field.|
| _Format_|Required| **Integer**| **VisFieldFormats**. The format for the new field.|

## Return value

Nothing


## Remarks

Using the  **AddField** method is similar to clicking **Field** on the **Insert** tab and inserting any of the following categories of fields in the text:


- Date/Time
    
- Document Info
    
- Geometry
    
- Object Info
    
- Page Info
    


To add a custom formula field, use the  **AddCustomField** method.

To specify language and calendar versions for Date/Time fields, use the  **AddFieldEx** method.

Constant values for  _Category_,  _Code_, and  _Format_ are declared by the Visio type library in **[VisFieldCategories](Visio.visfieldcategories.md)**, **[VisFieldCodes](Visio.visfieldcodes.md)**, and **[VisFieldFormats](Visio.visfieldformats.md)** respectively.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]