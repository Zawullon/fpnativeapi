{****************************************************************************
                        IMemoryManager.h
****************************************************************************}

type

  { IMemoryManager }
  P1CMemoryManager = ^T1CMemoryManager;
  T1CMemoryManager = record
    __vfptr: ^T1CMemoryManagerVTable;
  end;
  T1CMemoryManagerVTable = record
    _Destructor: procedure(This: P1CMemoryManager); cdecl;
    AllocMemory: function(This: P1CMemoryManager; var pMemory: Pointer; ulCountByte: culong): cbool; stdcall;
    FreeMemory: procedure(This: P1CMemoryManager; var pMemory: Pointer); stdcall;
  end;

