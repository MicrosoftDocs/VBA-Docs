---
title: Deftype statements must precede declarations
keywords: vblr6.chm1040057
f1_keywords:
- vblr6.chm1040057
ms.prod: office
ms.assetid: 1cbcf2e1-cd74-7d92-2d7a-2b6c3086e89a
ms.date: 06/08/2017
localization_priority: Normal
---


# Deftype statements must precede declarations

 **Def**_type_ statements include **DefInt**, **DefDbl**, **DefCur**, and so on. This error has the following causes and solutions:



- A [variable](../../Glossary/vbe-glossary.md#variable) [declaration](../../Glossary/vbe-glossary.md#declaration) precedes a **Def**_type_ statement at [module level](../../Glossary/vbe-glossary.md#module-level).
    
    Move the  **Def**_type_ statement to precede all variable declarations.
    
- A  **Def**_type_ statement appears in a [procedure](../../Glossary/vbe-glossary.md#procedure).
    
    Move the  **Def**_type_ statement to module level, preceding all variable declarations.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]