---
title: Command (Visual C++ syntax index with import)
ms.prod: access
ms.assetid: 9c3763f1-6242-a69c-bc2a-9d885f2b122a
ms.date: 10/12/2018
localization_priority: Normal
---


# Command (Visual C++ syntax index with #import)

**Applies to:** Access 2013 | Access 2016

## Methods

```csharp
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthadocancel_HV10294125.xml( ); 
 
_RecordsetPtr Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthcmdexecute_HV10294344.xml( VARIANT * RecordsAffected , VARIANT 
 * Parameters , long Options ); 
 
_ParameterPtr Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthcreateparam_HV10294243.xml( _bstr_t Name , enum 
 DataTypeEnum Type , enum ParameterDirectionEnum Direction , long Size , 
 const _variant_t & Value  =vtMissing); 

```

## Properties

```cs
_ConnectionPtr GetActiveConnection( ); 
void PutRefActiveConnection( struct _Connection * ppvObject ); 
void PutActiveConnection( const _variant_t & ppvObject ); 
__declspec(property(get=GetActiveConnection,put=PutRefActiveConnection)) 
 _ConnectionPtr Invalid DDUE based on source, error:link not allowed in code, link filename:mdproactivecon_HV10293988.xml; 
 
_bstr_t GetCommandText( ); 
void PutCommandText( _bstr_t pbstr ); 
__declspec(property(get=GetCommandText,put=PutCommandText)) _bstr_t 
 Invalid DDUE based on source, error:link not allowed in code, link filename:mdprocommandtext_HV10294195.xml; 
 
long GetCommandTimeout( ); 
void PutCommandTimeout( long pl ); 
__declspec(property(get=GetCommandTimeout,put=PutCommandTimeout)) long 
 Invalid DDUE based on source, error:link not allowed in code, link filename:mdprocommandtimeout_HV10294196.xml; 
 
void PutCommandType( enum CommandTypeEnum plCmdType ); 
enum CommandTypeEnum GetCommandType( ); 
__declspec(property(get=GetCommandType,put=PutCommandType)) enum 
 CommandTypeEnum Invalid DDUE based on source, error:link not allowed in code, link filename:mdprocommandtype_HV10294197.xml; 
 
VARIANT_BOOL GetPrepared( ); 
void PutPrepared( VARIANT_BOOL pfPrepared ); 
__declspec(property(get=GetPrepared,put=PutPrepared)) VARIANT_BOOL 
 Invalid DDUE based on source, error:link not allowed in code, link filename:mdproprepared_HV10294617.xml; 
 
ParametersPtr GetParameters( ); 
__declspec(property(get=GetParameters)) ParametersPtr 
 Invalid DDUE based on source, error:link not allowed in code, link filename:mdcolparameters_HV10294594.xml; 
 
_bstr_t GetName( ); 
void PutName( _bstr_t pbstrName ); 
__declspec(property(get=GetName,put=PutName)) _bstr_t Invalid DDUE based on source, error:link not allowed in code, link filename:mdproname_HV10294535.xml; 
 
long GetState( ); 
__declspec(property(get=GetState)) long Invalid DDUE based on source, error:link not allowed in code, link filename:mdprostate_HV10294804.xml; 

```

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]