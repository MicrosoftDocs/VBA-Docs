---
title: PtrSafe keyword (VBA)
ms.prod: office
ms.assetid: f413edb2-2839-efec-534a-eceb8d3da787
ms.date: 12/03/2018
localization_priority: Normal
---


# PtrSafe keyword

The **PtrSafe** keyword is used in this context: [Declare statement](declare-statement.md).

**Declare** statements with the **PtrSafe** keyword is the recommended syntax. Declare statements that include **PtrSafe** work correctly in the VBA7 development environment on both 32-bit and 64-bit platforms only after all data types in the **Declare** statement (parameters and return values) that need to store 64-bit quantities are updated to use [LongLong](longlong-data-type.md) for 64-bit integrals or [LongPtr](longptr-data-type.md) for pointers and handles. 

To ensure backwards compatibility with VBA version 6 and earlier, use the following construct:

```vb
#If VBA7 Then 
Declare PtrSafe Sub... 
#Else 
Declare Sub... 
#EndIf
```

When running in 64-bit versions of Office, **Declare** statements must include the **PtrSafe** keyword. The **PtrSafe** keyword asserts that a **Declare** statement is safe to run in 64-bit development environments.

Adding the **PtrSafe** keyword to a **Declare** statement only signifies that the **Declare** statement explicitly targets 64-bits. All data types within the statement that need to store 64-bits (including return values and parameters) must still be modified to hold 64-bit quantities by using either [LongLong](longlong-data-type.md) for 64-bit integrals or [LongPtr](longptr-data-type.md) for pointers and handles.

## See also

- [Keywords (VBA)](../keywords-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
