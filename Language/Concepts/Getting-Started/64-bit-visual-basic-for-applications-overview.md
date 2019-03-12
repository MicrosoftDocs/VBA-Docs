---
title: 64-bit Visual Basic for Applications overview
ms.prod: office
ms.assetid: a44e016f-1019-300e-5150-916ff32f70c1
ms.date: 12/21/2018
localization_priority: Priority
---


# 64-bit Visual Basic for Applications overview

Microsoft Visual Basic for Applications (VBA) is the version of Visual Basic that ships with Microsoft Office. In Microsoft Office 2010, VBA includes language features that enable VBA code to run correctly in both 32-bit and 64-bit environments.

> [!NOTE] 
> By default, Office 2010 installs the 32-bit version. You must explicitly choose to install the 64-bit version during setup.

Running VBA code that was written before the Office 2010 release (VBA version 6 and earlier) on a 64-bit platform can result in errors if the code is not modified to run in 64-bit versions of Office. Errors will result because VBA version 6 and earlier implicitly targets 32-bit platforms, and typically contains **[Declare statements](../../reference/user-interface-help/declare-statement.md)** that call into the Windows API by using 32-bit data types for pointers and handles. Because VBA version 6 and earlier does not have a specific data type for pointers or handles, it uses the **Long** data type, which is a 32-bit 4-byte data type, to reference pointers and handles. Pointers and handles in 64-bit environments are 8-byte 64-bit quantities. These 64-bit quantities cannot be held in 32-bit data types.

> [!NOTE] 
> You only need to modify VBA code if it runs in the 64-bit version of Microsoft Office.

The problem with running legacy VBA code in 64-bit Office is that trying to load 64-bits into a 32-bit data type truncates the 64-bit quantity. This can result in memory overruns, unexpected results in your code, and possible application failure.

To address this problem and enable VBA code to work correctly in both 32-bit and 64-bit environments, several language features have been added to VBA. The [table at the bottom of this document](#summary-of-vba7-language-updates) summarizes the new VBA language features. Three important additions are the **LongPtr** type alias, the **LongLong** data type, and the **PtrSafe** keyword.

- **[LongPtr](../../reference/user-interface-help/longptr-data-type.md)**. VBA now includes the variable type alias **LongPtr**. The actual data type that **LongPtr** resolves to depends on the version of Office that it is running in; **LongPtr** resolves to **Long** in 32-bit versions of Office, and **LongPtr** resolves to **LongLong** in 64-bit versions of Office. Use **LongPtr** for pointers and handles.
    
- **[LongLong](../../reference/user-interface-help/longlong-data-type.md)**. The **LongLong** data type is a signed 64-bit integer that is only available on 64-bit versions of Office. Use **LongLong** for 64-bit integrals. Conversion functions must be used to explicitly assign **LongLong** (including **LongPtr** on 64-bit platforms) to smaller integral types. Implicit conversions of **LongLong** to smaller integrals are not allowed.
    
- **[PtrSafe](../../reference/user-interface-help/ptrsafe-keyword.md)**. The **PtrSafe** keyword asserts that a **Declare** statement is safe to run in 64-bit versions of Office.
    
> [!IMPORTANT] 
> All **Declare** statements must now include the **PtrSafe** keyword when running in 64-bit versions of Office. It is important to understand that simply adding the **PtrSafe** keyword to a **Declare** statement only signifies that the **Declare** statement explicitly targets 64-bits. All data types within the statement that need to store 64-bits (including return values and parameters) must still be modified to hold 64-bit quantities.

> [!NOTE] 
> **Declare** statements with the **PtrSafe** keyword is the recommended syntax. **Declare** statements that include **PtrSafe** work correctly in the VBA7 development environment on both 32-bit and 64-bit platforms.
> 
> To ensure backwards compatibility in VBA7 and earlier use the following construct:
> 
> ```vb
>  #If VBA7 Then 
>  Declare PtrSafe Sub... 
>  #Else 
>  Declare Sub... 
>  #EndIf
> ```

Consider the following **Declare** statement examples. Running the unmodified **Declare** statement in 64-bit versions of Office will result in an error indicating that the **Declare** statement does not include the **PtrSafe** qualifier. The modified VBA example contains the **PtrSafe** qualifier, but notice that the return value (a pointer to the active window) returns a **Long** data type. On 64-bit Office, this is incorrect because the pointer needs to be 64-bits. The **PtrSafe** qualifier tells the compiler that the **Declare** statement is targeting 64-bits, so the statement executes without error. But because the return value has not been updated to a 64-bit data type, the return value is truncated, resulting in an incorrect value returned.

Following is an unmodified legacy VBA **Declare** statement example:

```vb
Declare Function GetActiveWindow Lib "user32" () As Long
```

The following VBA **Declare** statement example is modified to include the **PtrSafe** qualifier but still use a 32-bit return value:

```vb
Declare PtrSafe Function GetActiveWindow Lib "user32" () As Long
```

To reiterate, you must modify the **Declare** statement to include the **PtrSafe** qualifier, and you must update any variables within the statement that need to hold 64-bit quantities so that the variables use 64-bit data types.

Following is a VBA **Declare** statement example that is modified to include the **PtrSafe** keyword and is updated to use the proper 64-bit (**LongPtr**) data type:

```vb
Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
```

In summary, for code to work in 64-bit versions of Office, you need to locate and modify all existing **Declare** statements to use the **PtrSafe** qualifier. You also need to locate and modify all data types within these **Declare** statements that reference handles or pointers to use the new 64-bit compatible **LongPtr** type alias, and types that need to hold 64-bit integrals with the new **LongLong** data type. Additionally, you must update any user defined types (UDTs) that contain pointers or handles and 64-bit integrals to use 64-bit data types, and verify that all variable assignments are correct to prevent type mismatch errors.

## Writing code that works on both 32-bit and 64-bit Office

To write code that can port between both 32-bit and 64-bit versions of Office, you only need to use the new **LongPtr** type alias instead of **Long** or **LongLong** for all pointers and handle values. The **LongPtr** type alias will resolve to the correct **Long** or **LongLong** data type depending on which version of Office is running. 

Note that if you require different logic to execute, for example, you need to manipulate 64-bit values in large Excel projects, you can use the **Win64** conditional compilation constant as shown in the following section.


## Writing code that works on both Office 2010 (32-bit or 64-bit) and previous versions of Office

To write code that can work in both new and older versions of Office, you can use a combination of the new **VBA7** and **Win64** conditional [Compiler constants](compiler-constants.md). The **Vba7** conditional compiler constant is used to determine if code is running in version 7 of the VB editor (the VBA version that ships in Office 2010). The **Win64** conditional compiler constant is used to determine which version (32-bit or 64-bit) of Office is running.

```vb
#if Vba7 then 
'  Code is running in the new VBA7 editor 
     #if Win64 then 
     '  Code is running in 64-bit version of Microsoft Office 
     #else 
     '  Code is running in 32-bit version of Microsoft Office 
     #end if 
#else 
' Code is running in VBA version 6 or earlier 
#end if 
 
#If Vba7 Then 
Declare PtrSafe Sub... 
#Else 
Declare Sub... 
#EndIf 

```


## Summary of VBA7 language updates

The following table summarizes the new VBA language additions and provides an explanation of each.

|Name|Type|Description|
|:-----|:-----|:-----|
|**PtrSafe**|Keyword|Asserts that a **Declare** statement is targeted for 64-bit systems. Required on 64-bits.|
|**LongPtr**|Data type|Type alias that maps to **Long** on 32-bit systems, or **LongLong** on 64-bit systems.|
|**LongLong**|Data type|8-byte data type that is only available on 64-bit systems. Numeric type. Integer numbers in the range of -9,223,372,036,854,775,808 to 9,223,372,036,854,775,807.<br/><br/>**LongLong** is a valid declared type only on 64-bit platforms. Additionally, **LongLong** may not be implicitly converted to a smaller type (for example, you can't assign a **LongLong** to a **Long**). This is done to prevent inadvertent pointer truncation.<br/><br/>Explicit coercions are allowed, so in the previous example, you could apply **CLng** to a **LongLong** and assign the result to a **Long** (valid on 64-bit platforms only).|
|**^**|**LongLong** type-declaration character|Explicitly declares a literal value as a **LongLong**. Required to declare a **LongLong** literal that is larger than the maximum **Long** value (otherwise it will get implicitly converted to **double**).|
|**[CLngPtr](type-conversion-functions.md)**|type conversion function|Converts a simple expression to a **LongPtr**.|
|**[CLngLng](type-conversion-functions.md)**|type conversion function|Converts a simple expression to a **LongLong** data type (valid on 64-bit platforms only).|
|**[vbLongLong](vartype-constants.md)**|VarType constant|**LongLong** integer (valid on 64-bit platforms only).|
|**[DefLngPtr](deftype-statements.md)**|DefType statement|Sets the default data type for a range of variables as **LongPtr**.|
|**[DefLngLng](deftype-statements.md)**|DefType statement|Sets the default data type for a range of variables as **LongLong**.|

## See also

- [Visual Basic conceptual topics](../../reference/user-interface-help/visual-basic-conceptual-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
