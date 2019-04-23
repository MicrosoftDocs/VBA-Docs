---
title: Continue code execution (VBA)
keywords: vbhw6.chm1008878
f1_keywords:
- vbhw6.chm1008878
ms.prod: office
ms.assetid: 61035245-f12f-dea4-fa8e-5904f34d1bf3
ms.date: 12/27/2018
localization_priority: Normal
---


# Continue code execution

When you run your code, execution may stop if:

- An untrapped [run-time error](../Glossary/vbe-glossary.md#run-time-error) occurs.
    
- A trapped run-time error occurs, and **Break on All Errors** is selected on the **[General](../reference/user-interface-help/options-dialog-box.md#general-tab)** tab of the **Options** dialog box (**Tools** menu).
    
- A previously set [breakpoint](../Glossary/vbe-glossary.md#breakpoint) is encountered.
    
- A **[Stop](../reference/user-interface-help/stop-statement.md)** statement in your code is encountered, switching the mode to [break mode](../Glossary/vbe-glossary.md#break-mode).
    
- An **[End](../reference/user-interface-help/end-statement.md)** statement in your code is encountered, switching the mode to [design time](../Glossary/vbe-glossary.md#design-time).
    
- You halt execution manually at a given point.
    
- A [watch expression](../Glossary/vbe-glossary.md#watch-expression), which you set to break when the value has changed or break when the value is true, is encountered.
    
**To halt execution manually**

1. To switch to break mode, choose **Break** (CTRL+BREAK) from the **[Run](../reference/user-interface-help/run-menu.md)** menu, or use the toolbar shortcut: ![Toolbar button](../../images/tbr_brk_ZA01201682.gif).
    
2. To switch to design time, choose **Reset <projectname&gt;** from the **Run** menu, or use the toolbar shortcut: ![Toolbar button](../../images/tbr_end_ZA01201701.gif).
    

**To continue execution when your application has halted**

- On the **Run** menu, choose **Continue** (F5), or use the toolbar shortcut: ![Toolbar button](../../images/tbr_strt_ZA01201751.gif), or...
    
- On the **[Debug](../reference/user-interface-help/debug-menu.md)** menu, choose **Step Into** (F8), **Step Over** (SHIFT+F8), **Step Out** (CTRL+SHIFT+F8), or **Run To Cursor** (CTRL+F8).
    

**To continue execution when your application has halted because of a handled error**

- Press ALT+F8 to step through the error-handler, or...
    
- Press ALT+F5 to resume execution by running through the error-handler.
    

## See also

- [Visual Basic how-to topics](../reference/user-interface-help/visual-basic-how-to-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]