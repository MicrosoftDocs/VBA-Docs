---
title: Saved Property (VBA Add-In Object Model)
keywords: vbob6.chm1070966
f1_keywords:
- vbob6.chm1070966
ms.prod: office
ms.assetid: fd0e7762-5797-8fb2-03a8-b200c95cab19
ms.date: 06/08/2017
---


# Saved Property (VBA Add-In Object Model)



<<<<<<< HEAD
Returns a [Boolean](../../Glossary/vbe-glossary.md) value indicating whether or not the object was edited since the last time it was saved. Read/write.
=======
Returns a [Boolean](../../Glossary/vbe-glossary.md#boolean-data-type) value indicating whether or not the object was edited since the last time it was saved. Read/write.
>>>>>>> master
 **Return Values**
The  **Saved** property returns these values:


|**Value**|**Description**|
|:-----|:-----|
|**True**|The object has not been edited since the last time it was saved.|
|**False**|The object has been edited since the last time it was saved.|

<<<<<<< HEAD
 **Remarks**
=======
## Remarks

>>>>>>> master
The  **SaveAs** method sets the **Saved** property to **True**.

 **Note**  If you set the  **Saved** property to **False** in code, it returns **False**, and the object is marked as if it were edited since the last time it was saved.


