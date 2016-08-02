{****************************************************************************
                        Bindings
****************************************************************************}

{$PUSH}
{$MACRO ON}

{$define AddIn := PComponentBase(Pointer(This) - 0)^.Data}

function InitDoneBase_Init(This: PInitDoneBase; disp: Pointer): cbool; stdcall;
begin
  Result := False;
  if disp = nil then
    Exit;
  AddIn.FConnection := disp;
  Result := AddIn.Init();
end;

function InitDoneBase_setMemManager(This: PInitDoneBase; mem: Pointer): cbool; stdcall;
begin
  Result := False;
  if mem = nil then
    Exit;
  AddIn.FMemoryManager := mem;
  Result := True;
end;

function InitDoneBase_GetInfo(This: PInitDoneBase): clong; stdcall;
begin
  Result := AddIn.GetInfo;
end;

procedure InitDoneBase_Done(This: PInitDoneBase); stdcall;
begin
  AddIn.Done;
end;

{$define AddIn := PComponentBase(Pointer(This) - SizeOf(TInitDoneBase))^.Data}

function LanguageExtenderBase_RegisterExtensionAs(This: PLanguageExtenderBase; var wsExtensionName: PWideChar): cbool; stdcall;
var
  ExtensionName: String = '';
begin
  Result := False;
  if AddIn.FMemoryManager = nil then
    Exit;
  Result := AddIn.RegisterExtensionAs(ExtensionName);
  if Result then
    wsExtensionName := StringToWideChar(ExtensionName, AddIn.FMemoryManager);
end;

function LanguageExtenderBase_GetNProps(This: PLanguageExtenderBase): clong; stdcall;
begin
  Result := AddIn.GetNProps;
end;

function LanguageExtenderBase_FindProp(This: PLanguageExtenderBase; const wsPropName: PWideChar): clong; stdcall;
begin
  Result := AddIn.FindProp(WideCharToString(wsPropName));
end;

function LanguageExtenderBase_GetPropName(This: PLanguageExtenderBase; lPropNum: clong; lPropAlias: clong): PWideChar; stdcall;
var
  PropName: String;
begin
  Result := nil;
  if AddIn.FMemoryManager = nil then
    Exit;
  PropName := AddIn.GetPropName(lPropNum, lPropAlias);
  if PropName = '' then
    Exit;
  Result := StringToWideChar(PropName, AddIn.FMemoryManager);
end;

function LanguageExtenderBase_GetPropVal(This: PLanguageExtenderBase; const lPropNum: clong; pvarPropVal: P1CVariant): cbool; stdcall;
var
  Value: Variant;
begin
  Result := False;
  if AddIn.FMemoryManager = nil then
    Exit;
  Value := Unassigned;
  Result := AddIn.GetPropVal(lPropNum, Value);
  if Result then
    Result := To1CVariant(Value, pvarPropVal^, AddIn.FMemoryManager);
end;

function LanguageExtenderBase_SetPropVal(This: PLanguageExtenderBase; const lPropNum: clong; varPropVal: P1CVariant): cbool; stdcall;
var
  Value: Variant;
begin
  Value := Unassigned;
  Result := From1CVariant(varPropVal^, Value);
  if Result then
    Result := AddIn.SetPropVal(lPropNum, Value);
end;

function LanguageExtenderBase_IsPropReadable(This: PLanguageExtenderBase; const lPropNum: clong): cbool; stdcall;
begin
  Result := AddIn.IsPropReadable(lPropNum);
end;

function LanguageExtenderBase_IsPropWritable(This: PLanguageExtenderBase; const lPropNum: clong): cbool; stdcall;
begin
  Result := AddIn.IsPropWritable(lPropNum);
end;

function LanguageExtenderBase_GetNMethods(This: PLanguageExtenderBase): clong; stdcall;
begin
  Result := AddIn.GetNMethods;
end;

function LanguageExtenderBase_FindMethod(This: PLanguageExtenderBase; const wsMethodName: PWideChar): clong; stdcall;
begin
  Result := AddIn.FindMethod(WideCharToString(wsMethodName));
end;

function LanguageExtenderBase_GetMethodName(This: PLanguageExtenderBase; const lMethodNum: clong; const lMethodAlias: clong): PWideChar; stdcall;
var
  MethodName: String;
begin
  Result := nil;
  if AddIn.FMemoryManager = nil then
    Exit;
  MethodName := AddIn.GetMethodName(lMethodNum, lMethodAlias);
  if MethodName = '' then
    Exit;
  Result := StringToWideChar(MethodName, AddIn.FMemoryManager);
end;

function LanguageExtenderBase_GetNParams(This: PLanguageExtenderBase; const lMethodNum: clong): clong; stdcall;
begin
  Result := AddIn.GetNParams(lMethodNum);
end;

function LanguageExtenderBase_GetParamDefValue(This: PLanguageExtenderBase; const lMethodNum: clong; const lParamNum: clong; pvarParamDefValue: P1CVariant): cbool; stdcall;
var
  Value: Variant;
begin
  Result := False;
  if AddIn.FMemoryManager = nil then
    Exit;
  Value := Unassigned;
  Result := AddIn.GetParamDefValue(lMethodNum, lParamNum, Value);
  if Result then
    Result := To1CVariant(Value, pvarParamDefValue^, AddIn.FMemoryManager)
  else
    To1CVariant(Unassigned, pvarParamDefValue^, AddIn.FMemoryManager);
end;

function LanguageExtenderBase_HasRetVal(This: PLanguageExtenderBase; const lMethodNum: clong): cbool; stdcall;
begin
  Result := AddIn.HasRetVal(lMethodNum);
end;

function LanguageExtenderBase_CallAsProc(This: PLanguageExtenderBase; const lMethodNum: clong; paParams: P1CVariant; const lSizeArray: clong): cbool; stdcall;
var
  I: Integer;
  Params: array of Variant;
begin
  SetLength(Params, lSizeArray);
  for I := 0 to lSizeArray - 1 do
    Params[I] := Unassigned;
  for I := 0 to lSizeArray - 1 do
    begin
      Result := From1CVariant(paParams[I], Params[I]);
      if not Result then
        Break;
    end;
  if Result then
    Result := AddIn.CallAsProc(lMethodNum, Params);
end;

function LanguageExtenderBase_CallAsFunc(This: PLanguageExtenderBase; const lMethodNum: clong; pvarRetValue: P1CVariant; paParams: P1CVariant; const lSizeArray: clong): cbool; stdcall;
var
  I: Integer;
  Params: array of Variant;
  RetValue: Variant;
begin
  RetValue := Unassigned;
  SetLength(Params, lSizeArray);
  for I := 0 to lSizeArray - 1 do
    Params[I] := Unassigned;
  for I := 0 to lSizeArray - 1 do
    begin
      Result := From1CVariant(paParams[I], Params[I]);
      if not Result then
        Break;
    end;
  if Result then
    Result := AddIn.CallAsFunc(lMethodNum, RetValue, Params);
  if Result then
    Result := To1CVariant(RetValue, pvarRetValue^, AddIn.FMemoryManager)
  else
    To1CVariant(Unassigned, pvarRetValue^, AddIn.FMemoryManager);
end;

{$define AddIn := PComponentBase(Pointer(This) - (SizeOf(TInitDoneBase) + SizeOf(TLanguageExtenderBase)))^.Data}

procedure LocaleBase_SetLocale(This: PLocaleBase; const loc: PWideChar); stdcall;
begin
  AddIn.SetLocale(WideCharToString(loc));
end;

{$undef AddIn}

{$POP}

