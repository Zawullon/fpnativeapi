{****************************************************************************
                        AddInDefBase.h
****************************************************************************}

type
  PInterface = ^ TInterface;
  TInterface = record
  end;

  PAddInDefBase = ^TAddInDefBase;
  TAddInDefBase = record
    __vfptr: ^TAddInDefBaseVTable;
  end;
  TAddInDefBaseVTable = record
    _Destructor: procedure(This: PAddInDefBase); cdecl;
    AddError: function (This: PAddInDefBase; wcode: cushort; const source: PWideChar; const descr: PWideChar; scode: clong): cbool; stdcall;
    Read: function (This: PAddInDefBase; wszPropName: PWideChar; pVal: P1CVariant; pErrCode: pculong; errDescriptor: PPWideChar): cbool; stdcall;
    Write: function (This: PAddInDefBase; wszPropName: PWideChar; pVar: P1CVariant): cbool; stdcall;
    RegisterProfileAs: function (This: PAddInDefBase; wszProfileName: PWideChar): cbool; stdcall;
    SetEventBufferDepth: function (This: PAddInDefBase; lDepth: clong): cbool; stdcall;
    GetEventBufferDepth: function (This: PAddInDefBase): clong; stdcall;
    ExternalEvent: function (This: PAddInDefBase; wszSource: PWideChar; wszMessage: PWideChar; wszData: PWideChar): cbool; stdcall;
    CleanEventBuffer: procedure (This: PAddInDefBase); stdcall;
    SetStatusLine: function (This: PAddInDefBase; wszStatusLine: PWideChar): cbool; stdcall;
    ResetStatusLine: procedure (This: PAddInDefBase); stdcall;
  end;

  TInterfaces = (
    eIMsgBox = 0,
    eIPlatformInfo);

  PAddInDefBaseEx = ^TAddInDefBaseEx;
  TAddInDefBaseEx = record
    __vfptr: ^TAddInDefBaseExVTable;
  end;
  TAddInDefBaseExVTable = record
    Parent: TAddInDefBaseVTable;
    GetInterface: function (This: PAddInDefBaseEx; iface: TInterfaces): PInterface; stdcall;
  end;

  PAppInfo = ^TAppInfo;
  TAppInfo = record
    AppVersion: PWideChar;
    UserAgentInformation: PWideChar;
    Application: TAppType;
  end;

  PPlatformInfo = ^TPlatformInfo;
  TPlatformInfo = record
    __vfptr: ^TPlatformInfoVTable;
  end;
  TPlatformInfoVTable = record
    GetPlatformInfo: function (This: PPlatformInfo): PAppInfo; stdcall;
  end;

  PMsgBox = ^TMsgBox;
  TMsgBox = record
    __vfptr: ^TMsgBoxVTable;
  end;
  TMsgBoxVTable = record
    Confirm: function (This: PMsgBox; const queryText: PWideChar; retVal: P1CVariant): cbool; stdcall;
    Alert: function (This: PMsgBox; const text: PWideChar): cbool; stdcall;
  end;

