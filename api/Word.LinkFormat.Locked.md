---
title: LinkFormat.Locked property (Word)
keywords: vbawd10.chm154206221
f1_keywords:
- vbawd10.chm154206221
ms.prod: word
api_name:
- Word.LinkFormat.Locked
ms.assetid: 13125ef5-1809-f22e-abf6-d8781bc53e9a
ms.date: 06/08/2017
localization_priority: Normal
---


# LinkFormat.Locked property (Word)

 **True** if a **Field** , **InlineShape** , or **Shape** object is locked to prevent automatic updating. Read/write **Boolean**.


## Syntax

 _expression_. `Locked`

 _expression_ Required. A variable that represents a '[LinkFormat](Word.LinkFormat.md)' object.


## Remarks

If you use this property with a  **Shape** object that is a floating linked picture (a picture added with the **AddPicture** method of the **Shapes** object), an error occurs.


## See also


[LinkFormat Object](Word.LinkFormat.md)

