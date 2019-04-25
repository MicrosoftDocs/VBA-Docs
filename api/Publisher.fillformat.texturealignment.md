---
title: FillFormat.TextureAlignment property (Publisher)
keywords: vbapb10.chm2359575
f1_keywords:
- vbapb10.chm2359575
ms.prod: publisher
ms.assetid: 39fed9f2-f624-f978-3297-6b89a2dc3789
ms.date: 06/08/2017
localization_priority: Normal
---


# FillFormat.TextureAlignment property (Publisher)

Returns or sets the alignment (the origin of the coordinate grid) for the tiling of the texture fill. Read/write.


## Syntax

_expression_.**TextureAlignment**

_expression_ A variable that represents a **[FillFormat](publisher.fillformat.md)** object.


## Property value

 **MSOTEXTUREALIGNMENT**


## Remarks

The value returned by the  **TextureAlignment** property can be one of the following **MsoTextureAlignment** constants:


-  **msoTextureTopLeft**
    
-  **msoTextureTop**
    
-  **msoTextureTopRight**
    
-  **msoTextureLeft**
    
-  **msoTextureCenter**
    
-  **msoTextureRight**
    
-  **msoTextureBottomLeft**
    
-  **msoTextureBottom**
    
-  **msoTextureBottomRight**
    
The setting of the  **TextureAlignment** property corresponds to the **Alignment** setting on the **Fill** tab of the **Format Shape** dialog box in the Publisher 2013 user interface.


## See also


 [FillFormat Object](Publisher.FillFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]