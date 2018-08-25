---
title: Expected procedure, not user-defined type
keywords: vblr6.chm1035016
f1_keywords:
- vblr6.chm1035016
ms.prod: office
ms.assetid: c5fc855e-c844-792e-14fa-b861fa26ca84
ms.date: 06/08/2017
---


# Expected procedure, not user-defined type

There is no [procedure](../../Glossary/vbe-glossary.md#procedure) by this name in the current[scope](../../Glossary/vbe-glossary.md#scope), but there is a [user-defined type](../../Glossary/vbe-glossary.md#user-defined-type) by this name. You can call a procedure, but not a user-defined type. This error has the following cause and solution:



- The name of a user-defined type is used as a procedure call. Check the spelling of the procedure name, and make sure the procedure you are trying to call isn't private to another [module](../../Glossary/vbe-glossary.md#module).
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

