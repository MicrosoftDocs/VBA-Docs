---
title: SaveAs Method (VBA Add-In Object Model)
keywords: vbob6.chm102017
f1_keywords:
- vbob6.chm102017
ms.prod: office
ms.assetid: 622aa652-8093-be64-4128-9ad2c7fd1fe8
ms.date: 06/08/2017
---


# SaveAs Method (VBA Add-In Object Model)



Saves a project to a given location using a new filename.
<<<<<<< HEAD
 **Syntax**
 _object_**.SaveAs** **(**_newfilename_**As String)**
=======

## Syntax

_object_**.SaveAs** **(**_newfilename_**As String)**
>>>>>>> master
The  **SaveAs** method syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
<<<<<<< HEAD
| _object_|An [object expression](../../Glossary/vbe-glossary.md) that evaluates to an object in the Applies To list.|
| _newfilename_|Required. A [string expression](../../Glossary/vbe-glossary.md) specifying the new filename for the component to be saved.|

 **Remarks**
=======
| _object_|An [object expression](../../Glossary/vbe-glossary.md#object-expression) that evaluates to an object in the Applies To list.|
| _newfilename_|Required. A [string expression](../../Glossary/vbe-glossary.md#string-expression) specifying the new filename for the component to be saved.|

## Remarks

<<<<<<< HEAD
=======
## Remarks

>>>>>>> 54e0a75f224118db0d26fc9363ad519ad35ec788
>>>>>>> master
If a new path name is given, it is used. Otherwise, the old path name is used. If the new filename is invalid or refers to a read-only file, an error occurs.
The  **SaveAs** method can only be used on standalone projects. It generates a run-time error if you use it with a host project.

