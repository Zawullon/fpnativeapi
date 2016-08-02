{****************************************************************************
                        ComponentBase.h
****************************************************************************}

type

  { IInitDoneBase }
  PInitDoneBase = ^TInitDoneBase;
  TInitDoneBase = record
    __vfptr: ^TInitDoneBaseVTable;
  end;
  TInitDoneBaseVTable = record
    _Destructor: procedure(This: PInitDoneBase); cdecl;
    Init: function(This: PInitDoneBase; disp: Pointer): cbool; stdcall;
    setMemManager: function(This: PInitDoneBase; mem: Pointer): cbool; stdcall;
    GetInfo: function(This: PInitDoneBase): clong; stdcall;
    Done: procedure(This: PInitDoneBase); stdcall;
  end;

  { ILanguageExtenderBase }
  PLanguageExtenderBase = ^TLanguageExtenderBase;
  TLanguageExtenderBase = record
    __vfptr: ^TLanguageExtenderBaseVTable;
  end;
  TLanguageExtenderBaseVTable = record
    _Destructor: procedure(This: PLanguageExtenderBase); cdecl;
    RegisterExtensionAs: function(This: PLanguageExtenderBase; var wsExtensionName: PWideChar): cbool; stdcall;
    GetNProps: function(This: PLanguageExtenderBase): clong; stdcall;
    FindProp: function(This: PLanguageExtenderBase; const wsPropName: PWideChar): clong; stdcall;
    GetPropName: function(This: PLanguageExtenderBase; lPropNum: clong; lPropAlias: clong): PWideChar; stdcall;
    GetPropVal: function(This: PLanguageExtenderBase; const lPropNum: clong; pvarPropVal: P1CVariant): cbool; stdcall;
    SetPropVal: function(This: PLanguageExtenderBase; const lPropNum: clong; varPropVal: P1CVariant): cbool; stdcall;
    IsPropReadable: function(This: PLanguageExtenderBase; const lPropNum: clong): cbool; stdcall;
    IsPropWritable: function(This: PLanguageExtenderBase; const lPropNum: clong): cbool; stdcall;
    GetNMethods: function(This: PLanguageExtenderBase): clong; stdcall;
    FindMethod: function(This: PLanguageExtenderBase; const wsMethodName: PWideChar): clong; stdcall;
    GetMethodName: function(This: PLanguageExtenderBase; const lMethodNum: clong; const lMethodAlias: clong): PWideChar; stdcall;
    GetNParams: function(This: PLanguageExtenderBase; const lMethodNum: clong): clong; stdcall;
    GetParamDefValue: function(This: PLanguageExtenderBase; const lMethodNum: clong; const lParamNum: clong; pvarParamDefValue: P1CVariant): cbool; stdcall;
    HasRetVal: function(This: PLanguageExtenderBase; const lMethodNum: clong): cbool; stdcall;
    CallAsProc: function(This: PLanguageExtenderBase; const lMethodNum: clong; paParams: P1CVariant; const lSizeArray: clong): cbool; stdcall;
    CallAsFunc: function(This: PLanguageExtenderBase; const lMethodNum: clong; pvarRetValue: P1CVariant; paParams: P1CVariant; const lSizeArray: clong): cbool; stdcall;
  end;

  { ILocaleBase }
  PLocaleBase = ^TLocaleBase;
  TLocaleBase = record
    __vfptr: ^TLocaleBaseVTable;
  end;
  TLocaleBaseVTable = record
    _Destructor: procedure(This: PLocaleBase); cdecl;
    SetLocale: procedure(This: PLocaleBase; const loc: PWideChar); stdcall;
  end;

  { IComponentBase }
  PComponentBase = ^TComponentBase;
  TComponentBase = record
    InitDoneBase: TInitDoneBase;
    LanguageExtenderBase: TLanguageExtenderBase;
    LocaleBase: TLocaleBase;
    __vfptr: ^TComponentBaseVTable;
    { FPC object }
    Data: TAddIn;
    Factory: TAddInFactory;
  end;
  TComponentBaseVTable = record
    _Destructor: procedure(This: PComponentBase); cdecl;
  end;

