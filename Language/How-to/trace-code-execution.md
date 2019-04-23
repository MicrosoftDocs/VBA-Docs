---
title: Trace code execution (VBA)
keywords: vbhw6.chm1009047
f1_keywords:
- vbhw6.chm1009047
ms.prod: office
ms.assetid: 1eccc024-5317-3f74-2259-1a8c6dd5a785
ms.date: 12/27/2018
localization_priority: Normal
---


# Trace code execution

You trace code execution because it may not always be obvious which statement is executed first. Use these techniques to trace the execution of code:

- **Step Into**: Traces through each line of code and steps into [procedures](../Glossary/vbe-glossary.md#procedure). This allows you to view the effect of each statement on [variables](../Glossary/vbe-glossary.md#variable).
    
- **Step Over**: Executes each procedure as if it were a single statement. Use this instead of **Step Into** to step across procedure calls rather than into the called procedure.
    
- **Step Out**: Executes all remaining code in a procedure as if it were a single statement, and exits to the next statement in the procedure that caused the procedure to be called initially.
    
- **Run To Cursor**: Allows you to select a statement in your code where you want execution to stop. This allows you to "step over" sections of code, for example, large loops.
    
**To trace execution from the current statement**

- From the **[Debug](../reference/user-interface-help/debug-menu.md)** menu, choose **Step Into** (F8), **Step Over** (SHIFT+F8), **Step Out** (CTRL+SHIFT+F8), or **Run To Cursor** (CTRL+F8).
    
**To trace execution from the beginning of the program**

- From the **Debug** menu, choose **Step Into** (F8), **Step Over** (SHIFT+F8), **Step Out** (CTRL+SHIFT+F8), or **Run To Cursor** (CTRL+F8).
    
## See also

- [Visual Basic how-to topics](../reference/user-interface-help/visual-basic-how-to-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]