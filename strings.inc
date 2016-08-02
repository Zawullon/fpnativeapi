{****************************************************************************
                        Converting strings
****************************************************************************}

const
  EmptyWideChar: WideChar = #$0000;

function StringToWideChar(const Source: String; MemoryManager: P1CMemoryManager): PWideChar; inline;
var
  WS: WideString;
  ResultSize: SizeInt;
begin
  WS := UTF8Decode(Source);
  ResultSize := (Length(WS) + 1) * SizeOf(WideChar);
  if MemoryManager = nil then
    Result := GetMem(ResultSize)
  else
    begin
      Result := nil;
      if not MemoryManager^.__vfptr^.AllocMemory(MemoryManager, Result, ResultSize) then
        Result := nil;
    end;
  if Result = nil then
    Exit;
  if ResultSize = SizeOf(WideChar) then
    Move(EmptyWideChar, Result^, ResultSize)
  else
    Move(WS[1], Result^, ResultSize);
end;

function WideCharToString(Source: PWideChar): String; inline;
var
  WS: WideString;
begin
  WS := Source;
  Result := UTF8Encode(WS);
end;

