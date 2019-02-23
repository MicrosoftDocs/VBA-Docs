---
title: Execute a specific statement (VBA)
keywords: vbhw6.chm1105192
f1_keywords:
- vbhw6.chm1105192
ms.prod: office
ms.assetid: 126c4d53-47a9-b36c-ef7c-d246d3f1cd5d
ms.date: 12/27/2018
localization_priority: Normal
---


# Execute a specific statement

While execution of your code is halted, you can control the execution sequence of [statements](../Glossary/vbe-glossary.md#statement) within a [procedure](../Glossary/vbe-glossary.md#procedure). You can resume execution at a statement you choose without executing any intervening code.

**To set the next statement to be executed**

1. In the Code window, position the insertion point anywhere within the statement.
    
2. On the **[Debug](../reference/user-interface-help/debug-menu.md)** menu, choose **Set Next Statement** (CTRL+F9), or on Windows, position the mouse pointer in the [margin indicator](../Glossary/vbe-glossary.md#margin-indicator) next to the current execution point.
    
3. Drag the yellow arrow in the margin indicator to the statement that you want to execute next.
    
   > [!NOTE] 
   > You can only skip to statements within the same procedure.

Used in combination with **Step Into**, executing specific statements with the **Set Next Statement** command enables you to step through procedures one statement at a time, and to closely examine your code. It's also helpful for correcting or avoiding [run-time error](../Glossary/vbe-glossary.md#run-time-error) conditions.
    

## See also

- [Visual Basic how-to topics](../reference/user-interface-help/visual-basic-how-to-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]