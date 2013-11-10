Attribute VB_Name = "KernelStructs"
Public Type DISPATCHER_HEADER
   '+0x000 Type             : UChar
    Type As Byte
   '+0x001 Absolute         : UChar
    Absolute As Byte
   '+0x002 Size             : UChar
    Size As Byte
   '+0x003 Inserted         : UChar
    Inserted As Byte
   '+0x004 SignalState      : Int4B
    SignalState As Long
   '+0x008 WaitListHead     : _LIST_ENTRY
    WaitListHead As LIST_ENTRY
End Type

Public Type K_SEMAPHORE
    Fill(1 To &H14) As Byte
End Type

Public Type KGDENTRY
    Fill(1 To 2) As Long
End Type

Public Type KIDENTRY
    Fill(1 To 2) As Long
End Type

Public Type K_Process
    '+0x000 Header           : _DISPATCHER_HEADER
     Header As DISPATCHER_HEADER
    '+0x010 ProfileListHead  : _LIST_ENTRY
     ProfileListHead As LIST_ENTRY
    '+0x018 DirectoryTableBase : [2] Uint4B
     DirectoryTableBase(1 To 2) As Long
    '+0x020 LdtDescriptor    : _KGDTENTRY
     LdtDescriptor As KGDENTRY
    '+0x028 Int21Descriptor  : _KIDTENTRY
     Int21Descriptor As KIDENTRY
    '+0x030 IopmOffset       : Uint2B
     IopmOffset As Integer
    '+0x032 Iopl             : UChar
     Iopl As Byte
    '+0x033 Unused           : UChar
     Unused As Byte
    '+0x034 ActiveProcessors : Uint4B
     ActiveProcessors As Long
    '+0x038 KernelTime       : Uint4B
     KernelTime As Long
    '+0x03c UserTime         : Uint4B
     UserTime As Long
    '+0x040 ReadyListHead    : _LIST_ENTRY
     ReadyListHead As LIST_ENTRY
    '+0x048 SwapListEntry    : _SINGLE_LIST_ENTRY
     SwapListEntry As Long
    '+0x04c VdmTrapcHandler  : Ptr32 Void
     VdmTrapcHandler As Long
    '+0x050 ThreadListHead   : _LIST_ENTRY
     ThreadListHead As LIST_ENTRY
    '+0x058 ProcessLock      : Uint4B
     ProcessLock As Long
    '+0x05c Affinity         : Uint4B
     Affinity As Long
    '+0x060 StackCount       : Uint2B
     StackCount As Integer
    '+0x062 BasePriority     : Char
     BasePriority As Byte
    '+0x063 ThreadQuantum    : Char
     ThreadQuantum As Byte
    '+0x064 AutoAlignment    : UChar
     AutoAlignment As Byte
    '+0x065 State            : UChar
     State As Byte
    '+0x066 ThreadSeed       : UChar
     ThreadSeed As Byte
    '+0x067 DisableBoost     : UChar
     DisableBoost As Byte
    '+0x068 PowerState       : UChar
     PowerState As Byte
    '+0x069 DisableQuantum   : UChar
     DisableQuantum As Byte
    '+0x06a IdealNode        : UChar
     IdealNode As Byte
    '+0x06b Flags            : _KEXECUTE_OPTIONS  ?
    '+0x06b ExecuteOptions   : UChar
     ExecuteOptions As Byte
End Type

Public Type MMSUPPORT
    Fill(1 To &H40) As Byte
End Type

Public Type E_Process
    '+0x000 Pcb              : _KPROCESS
    Pcb As K_Process
    '+0x06c ProcessLock      : _EX_PUSH_LOCK
    ProcessLock As Long
    '+0x070 CreateTime       : _LARGE_INTEGER
    CreateTime As FILETIME
    '+0x078 ExitTime         : _LARGE_INTEGER
    ExitTime As FILETIME
    '+0x080 RundownProtect   : _EX_RUNDOWN_REF
    RundownProtect As Long
    '+0x084 UniqueProcessId  : Ptr32 Void
    UniqueProcessId As Long
    '+0x088 ActiveProcessLinks : _LIST_ENTRY
    ActiveProcessLinks As LIST_ENTRY
    '+0x090 QuotaUsage       : [3] Uint4B
    QuotaUsage(1 To 3) As Long
    '+0x09c QuotaPeak        : [3] Uint4B
    QuotaPeak(1 To 3) As Long
    '+0x0a8 CommitCharge     : Uint4B
    CommitCharge As Long
    '+0x0ac PeakVirtualSize  : Uint4B
    PeakVirtualSize As Long
    '+0x0b0 VirtualSize      : Uint4B
    VirtualSize As Long
    '+0x0b4 SessionProcessLinks : _LIST_ENTRY
    SessionProcessLinks As LIST_ENTRY
    '+0x0bc DebugPort        : Ptr32 Void
    DebugPort As Long
    '+0x0c0 ExceptionPort    : Ptr32 Void
    ExceptionPort As Long
    '+0x0c4 ObjectTable      : Ptr32 _HANDLE_TABLE
    ObjectTable As Long
    '+0x0c8 Token            : _EX_FAST_REF
    Token As Long
    '+0x0cc WorkingSetLock   : _FAST_MUTEX
    WorkingSetLock(1 To &H8) As Long
    '+0x0ec WorkingSetPage   : Uint4B
    WorkingSetPage As Long
    '+0x0f0 AddressCreationLock : _FAST_MUTEX
    AddressCreationLock(1 To &H8) As Long
    '+0x110 HyperSpaceLock   : Uint4B
    HyperSpaceLock As Long
    '+0x114 ForkInProgress   : Ptr32 _ETHREAD
    ForkInProgress As Long
    '+0x118 HardwareTrigger  : Uint4B
    HardwareTrigger As Long
    '+0x11c VadRoot          : Ptr32 Void
    VadRoot As Long
    '+0x120 VadHint          : Ptr32 Void
    VadHint As Long
    '+0x124 CloneRoot        : Ptr32 Void
    CloneRoot As Long
    '+0x128 NumberOfPrivatePages : Uint4B
    NumberOfPrivatePages As Long
    '+0x12c NumberOfLockedPages : Uint4B
    NumberOfLockedPages As Long
    '+0x130 Win32Process     : Ptr32 Void
    Win32Process As Long
    '+0x134 Job              : Ptr32 _EJOB
    Job As Long
    '+0x138 SectionObject    : Ptr32 Void
    SectionObject As Long
    '+0x13c SectionBaseAddress : Ptr32 Void
    SectionBaseAddress As Long
    '+0x140 QuotaBlock       : Ptr32 _EPROCESS_QUOTA_BLOCK
    QuotaBlock As Long
    '+0x144 WorkingSetWatch  : Ptr32 _PAGEFAULT_HISTORY
    WorkingSetWatch As Long
    '+0x148 Win32WindowStation : Ptr32 Void
    Win32WindowStation As Long
    '+0x14c InheritedFromUniqueProcessId : Ptr32 Void
    InheritedFromUniqueProcessId As Long
    '+0x150 LdtInformation   : Ptr32 Void
    LdtInformation As Long
    '+0x154 VadFreeHint      : Ptr32 Void
    BadFreeHint As Long
    '+0x158 VdmObjects       : Ptr32 Void
    VdmObjects As Long
    '+0x15c DeviceMap        : Ptr32 Void
    DeviceMap As Long
    '+0x160 PhysicalVadList  : _LIST_ENTRY
    PhysicalVadList As LIST_ENTRY
    '+0x168 PageDirectoryPte : _HARDWARE_PTE  ?
    '+0x168 Filler           : Uint8B
    Filler As FILETIME
    '+0x170 Session          : Ptr32 Void
    Session As Long
    '+0x174 ImageFileName    : [16] UChar
    'ImageFileName As String * 16
    ImageFileName(1 To 16) As Byte
    '+0x184 JobLinks         : _LIST_ENTRY
    JobLinks As LIST_ENTRY
    '+0x18c LockedPagesList  : Ptr32 Void
    LockedPagesList As Long
    '+0x190 ThreadListHead   : _LIST_ENTRY
    ThreadListHead As LIST_ENTRY
    '+0x198 SecurityPort     : Ptr32 Void
    SecurityPort As Long
    '+0x19c PaeTop           : Ptr32 Void
    PaeTop As Long
    '+0x1a0 ActiveThreads    : Uint4B
    ActiveThreads As Long
    '+0x1a4 GrantedAccess    : Uint4B
    GrantedAccess As Long
    '+0x1a8 DefaultHardErrorProcessing : Uint4B
    DefaultHardErrorProcessing As Long
    '+0x1ac LastThreadExitStatus : Int4B
    LastThreadExitStatus As Long
    '+0x1b0 Peb              : Ptr32 _PEB
    Peb As Long
    '+0x1b4 PrefetchTrace    : _EX_FAST_REF
    PrefetchTrace As Long
    '+0x1b8 ReadOperationCount : _LARGE_INTEGER
    ReadOperationCount As FILETIME
    '+0x1c0 WriteOperationCount : _LARGE_INTEGER
    WriteOperationCount As FILETIME
    '+0x1c8 OtherOperationCount : _LARGE_INTEGER
    OtherOperationCount As FILETIME
    '+0x1d0 ReadTransferCount : _LARGE_INTEGER
    ReadTransferCount As FILETIME
    '+0x1d8 WriteTransferCount : _LARGE_INTEGER
    WriteTransferCount As FILETIME
    '+0x1e0 OtherTransferCount : _LARGE_INTEGER
    OtherTransferCount As FILETIME
    '+0x1e8 CommitChargeLimit : Uint4B
    CommitChargeLimit As Long
    '+0x1ec CommitChargePeak : Uint4B
    CommitChargePeak As Long
    '+0x1f0 AweInfo          : Ptr32 Void
    AweInfo As Long
    '+0x1f4 SeAuditProcessCreationInfo : _SE_AUDIT_PROCESS_CREATION_INFO
    SeAuditProcessCreationInfo As Long
    '+0x1f8 Vm               : _MMSUPPORT
    Vm As MMSUPPORT
    '+0x238 LastFaultCount   : Uint4B
    LastFaultCount As Long
    '+0x23c ModifiedPageCount : Uint4B
    ModifiedPageCount As Long
    '+0x240 NumberOfVads     : Uint4B
    NumberOfVads As Long
    '+0x244 JobStatus        : Uint4B
    JobStatus As Long
    '+0x248 Flags            : Uint4B
    '+0x248 CreateReported   : Pos 0, 1 Bit
    '+0x248 NoDebugInherit   : Pos 1, 1 Bit
    '+0x248 ProcessExiting   : Pos 2, 1 Bit
    '+0x248 ProcessDelete    : Pos 3, 1 Bit
    '+0x248 Wow64SplitPages  : Pos 4, 1 Bit
    '+0x248 VmDeleted        : Pos 5, 1 Bit
    '+0x248 OutswapEnabled   : Pos 6, 1 Bit
    '+0x248 Outswapped       : Pos 7, 1 Bit
    '+0x248 ForkFailed       : Pos 8, 1 Bit
    '+0x248 HasPhysicalVad   : Pos 9, 1 Bit
    '+0x248 AddressSpaceInitialized : Pos 10, 2 Bits
    '+0x248 SetTimerResolution : Pos 12, 1 Bit
    '+0x248 BreakOnTermination : Pos 13, 1 Bit
    '+0x248 SessionCreationUnderway : Pos 14, 1 Bit
    '+0x248 WriteWatch       : Pos 15, 1 Bit
    '+0x248 ProcessInSession : Pos 16, 1 Bit
    '+0x248 OverrideAddressSpace : Pos 17, 1 Bit
    '+0x248 HasAddressSpace  : Pos 18, 1 Bit
    '+0x248 LaunchPrefetched : Pos 19, 1 Bit
    '+0x248 InjectInpageErrors : Pos 20, 1 Bit
    '+0x248 VmTopDown        : Pos 21, 1 Bit
    '+0x248 Unused3          : Pos 22, 1 Bit
    '+0x248 Unused4          : Pos 23, 1 Bit
    '+0x248 VdmAllowed       : Pos 24, 1 Bit
    '+0x248 Unused           : Pos 25, 5 Bits
    '+0x248 Unused1          : Pos 30, 1 Bit
    '+0x248 Unused2          : Pos 31, 1 Bit
    Flags As Long
    '+0x24c ExitStatus       : Int4B
    ExitStatus As Long
    '+0x250 NextPageColor    : Uint2B
    NextPageColor As Integer
    '+0x252 SubSystemMinorVersion : UChar
    SubSystemMinorVersion As Byte
    '+0x253 SubSystemMajorVersion : UChar
    SubSystemMajorVersion As Byte
    '+0x252 SubSystemVersion : Uint2B
    SubSystemVersion As Integer
    '+0x254 PriorityClass    : UChar
    PriorityClass As Byte
    '+0x255 WorkingSetAcquiredUnsafe : UChar
    WorkingSetAcquiredUnsafe As Byte
    '?
    a As Integer
    '+0x258 Cookie           : Uint4B
    Cookie As Long
End Type

'---------------- Thread
Public Type K_APC_STATE
    Fill(1 To &H18) As Byte
End Type

Public Type K_WAIT_BLOCK
   '+0x000 WaitListEntry    : _LIST_ENTRY
    WaitListEntry As LIST_ENTRY
   '+0x008 Thread           : Ptr32 _KTHREAD
    Thread As Long
   '+0x00c Object           : Ptr32 Void
    Object As Long
   '+0x010 NextWaitBlock    : Ptr32 _KWAIT_BLOCK
    NextWaitBlock As Long
   '+0x014 WaitKey          : Uint2B
    WaitKey As Integer
   '+0x016 WaitType         : Uint2B
   WaitType As Integer
End Type

Public Type K_TIMER
    Fill(1 To &H28) As Byte
End Type

Public Type K_APC
    Fill(1 To &H30) As Byte
End Type

Public Type K_Thread
     '+0x000 Header           : _DISPATCHER_HEADER
     Header As DISPATCHER_HEADER
     '+0x010 MutantListHead   : _LIST_ENTRY
     NutantListHead As LIST_ENTRY
    '+0x018 InitialStack     : Ptr32 Void
     InitialStack As Long
    '+0x01c StackLimit       : Ptr32 Void
     StackLimit As Long
    '+0x020 Teb              : Ptr32 Void
     Tab As Long
    '+0x024 TlsArray         : Ptr32 Void
     TlsArray As Long
    '+0x028 KernelStack      : Ptr32 Void
     KernelStack As Long
    '+0x02c DebugActive      : UChar
     DebugActive As Byte
    '+0x02d State            : UChar
     State As Byte
    '+0x02e Alerted          : [2] UChar
     Alerted(1 To 2) As Byte
    '+0x030 Iopl             : UChar
     Topl As Byte
    '+0x031 NpxState         : UChar
     NpxState As Byte
    '+0x032 Saturation       : Char
     Saturation As Byte
    '+0x033 Priority         : Char
     Priority As Byte
    '+0x034 ApcState         : _KAPC_STATE
     ApcState As K_APC_STATE
    '+0x04c ContextSwitches  : Uint4B
     ContextSwitches As Long
    '+0x050 IdleSwapBlock    : UChar
     IdleSwapBlock As Byte
    '+0x051 Spare0           : [3] UChar
     Spare0(1 To 3) As Byte
    '+0x054 WaitStatus       : Int4B
     WaitStatus As Long
    '+0x058 WaitIrql         : UChar
     WaitIrql As Byte
    '+0x059 WaitMode         : Char
     WaitMode As Byte
    '+0x05a WaitNext         : UChar
     WaitNext As Byte
    '+0x05b WaitReason       : UChar
     WaitReason As Byte
    '+0x05c WaitBlockList    : Ptr32 _KWAIT_BLOCK
     WaitBlockList As Long
    '+0x060 WaitListEntry    : _LIST_ENTRY
     WaitListEntry As LIST_ENTRY
    '+0x060 SwapListEntry    : _SINGLE_LIST_ENTRY  ?
     'SwapListEntry As Long
    '+0x068 WaitTime         : Uint4B
     WaitTime As Long
    '+0x06c BasePriority     : Char
     BasePriority As Byte
    '+0x06d DecrementCount   : UChar
     DecrementCount As Byte
    '+0x06e PriorityDecrement : Char
     PriorityDecrement As Byte
    '+0x06f Quantum          : Char
     Quantum As Byte
    '+0x070 WaitBlock        : [4] _KWAIT_BLOCK
     WaitBlock(1 To 4) As K_WAIT_BLOCK
    '+0x0d0 LegoData         : Ptr32 Void
     LegoData As Long
    '+0x0d4 KernelApcDisable : Uint4B
     KernelApcDisable As Long
    '+0x0d8 UserAffinity     : Uint4B
     UserAffinity As Long
    '+0x0dc SystemAffinityActive : UChar
     SystemAffinityActive As Byte
    '+0x0dd PowerState       : UChar
     PowerState As Byte
    '+0x0de NpxIrql          : UChar
     NpxIrql As Byte
    '+0x0df InitialNode      : UChar
     InitialNode As Byte
    '+0x0e0 ServiceTable     : Ptr32 Void             //GUI线程指向KeServiceDescriptorTableShadow，CUI线程指向KeServiceDescriptorTable
     ServiceTable As Long
    '+0x0e4 Queue            : Ptr32 _KQUEUE
     Queue As Long
    '+0x0e8 ApcQueueLock     : Uint4B
     ApcQueueLock As Long
     '(? E8+4=EC != F0)
     Fill As Long
     Fill2 As Long
    '+0x0f0 Timer            : _KTIMER
     Timer As K_TIMER
    '+0x118 QueueListEntry   : _LIST_ENTRY
     QueueListEntry As LIST_ENTRY
    '+0x120 SoftAffinity     : Uint4B
     SoftAffinity As Long
    '+0x124 Affinity         : Uint4B
     Affinity As Long
    '+0x128 Preempted        : UChar
     Preempted As Byte
    '+0x129 ProcessReadyQueue : UChar
     ProcessReadyQueue As Byte
    '+0x12a KernelStackResident : UChar
     KernelStackResident As Byte
    '+0x12b NextProcessor    : UChar
     NextProcessor As Byte
    '+0x12c CallbackStack    : Ptr32 Void
     CallbackStack As Long
    '+0x130 Win32Thread      : Ptr32 Void
     Win32Thread As Long
    '+0x134 TrapFrame        : Ptr32 _KTRAP_FRAME
     TrapFrame As Long
    '+0x138 ApcStatePointer  : [2] Ptr32 _KAPC_STATE
     AppStatePoint(1 To 2) As Long
    '+0x140 PreviousMode     : Char
     PreviousMode As Byte
    '+0x141 EnableStackSwap  : UChar
     EnableStackSwap As Byte
    '+0x142 LargeStack       : UChar
     LargeStack As Byte
    '+0x143 ResourceIndex    : UChar
     ResourceIndex As Byte
    '+0x144 KernelTime       : Uint4B
     KernelTime As Long
    '+0x148 UserTime         : Uint4B
     UserTime As Long
    '+0x14c SavedApcState    : _KAPC_STATE
     SavedApcState As K_APC_STATE
    '+0x164 Alertable        : UChar
     Alertable As Byte
    '+0x165 ApcStateIndex    : UChar
     ApcStateIndex As Byte
    '+0x166 ApcQueueable     : UChar
     ApcQueueable As Byte
    '+0x167 AutoAlignment    : UChar
     AutoAlignment As Byte
    '+0x168 StackBase        : Ptr32 Void
     StackBase As Long
    '+0x16c SuspendApc       : _KAPC
     SuspendApc As K_APC
    '+0x19c SuspendSemaphore : _KSEMAPHORE
     SuspendSemaphore As K_SEMAPHORE
    '+0x1b0 ThreadListEntry  : _LIST_ENTRY
     ThreadListEntry As LIST_ENTRY
    '+0x1b8 FreezeCount      : Char
     FreezeCount As Byte
    '+0x1b9 SuspendCount     : Char
     SuspendCount As Byte
    '+0x1ba IdealProcessor   : UChar
     IdealProcessor As Byte
    '+0x1bb DisableBoost     : UChar
     DisableBoost As Byte
End Type
Public Type E_Thread
   '+0x000 Tcb              : _KTHREAD
   Tcb As K_Thread
   '+0x1c0 CreateTime       : _LARGE_INTEGER
   '+0x1c0 NestedFaultCount : Pos 0, 2 Bits
   '+0x1c0 ApcNeeded        : Pos 2, 1 Bit
   CreateTime As FILETIME
   '+0x1c8 ExitTime         : _LARGE_INTEGER
   ExitTime As FILETIME
   '+0x1c8 LpcReplyChain    : _LIST_ENTRY   ?
   'LpcReplyChain As LIST_ENTRY
   '+0x1c8 KeyedWaitChain   : _LIST_ENTRY   ?
   'KeyedWaitChain As LIST_ENTRY
   '+0x1d0 ExitStatus       : Int4B
   ExitStatus As Long
   '+0x1d0 OfsChain         : Ptr32 Void    ?
   'OfsChain As Long
   '+0x1d4 PostBlockList    : _LIST_ENTRY
   PostBlockList As LIST_ENTRY
   '+0x1dc TerminationPort  : Ptr32 _TERMINATION_PORT
   TerminationPort As Long
   '+0x1dc ReaperLink       : Ptr32 _ETHREAD   ?
   'ReaperLink As Long
   '+0x1dc KeyedWaitValue   : Ptr32 Void       ?
   'KeyedWaitValue As Long
   '+0x1e0 ActiveTimerListLock : Uint4B
   ActiveTimerListLock As Long
   '+0x1e4 ActiveTimerListHead : _LIST_ENTRY
   ActiveTimerListHead As LIST_ENTRY
   '+0x1ec Cid              : _CLIENT_ID
   Cid As CLIENT_ID
   '+0x1f4 LpcReplySemaphore : _KSEMAPHORE
   LpcReplySemaphore As K_SEMAPHORE
   '+0x1f4 KeyedWaitSemaphore : _KSEMAPHORE  ?
   '+0x208 LpcReplyMessage  : Ptr32 Void
   LpcReplyMessage As Long
   '+0x208 LpcWaitingOnPort : Ptr32 Void
   'LpcWaitingOnPort As Long                 ?
   '+0x20c ImpersonationInfo : Ptr32 _PS_IMPERSONATION_INFORMATION
   ImpersonationInfo As Long
   '+0x210 IrpList          : _LIST_ENTRY
   IrpList As LIST_ENTRY
   '+0x218 TopLevelIrp      : Uint4B
   TopLevelIrp As Long
   '+0x21c DeviceToVerify   : Ptr32 _DEVICE_OBJECT
   DeviceToVerify As Long
   '+0x220 ThreadsProcess   : Ptr32 _EPROCESS
   ThreadsProcess As Long
   '+0x224 StartAddress     : Ptr32 Void
   StartAddress As Long
   '+0x228 Win32StartAddress : Ptr32 Void
   Win32StartAddress As Long
   '+0x228 LpcReceivedMessageId : Uint4B
   'LpcReceivedMessageId As Long            ?
   '+0x22c ThreadListEntry  : _LIST_ENTRY
   ThreadListEntry As LIST_ENTRY
   '+0x234 RundownProtect   : _EX_RUNDOWN_REF
   RundownProtect As Long
   '+0x238 ThreadLock       : _EX_PUSH_LOCK
   ThreadLock As Long
   '+0x23c LpcReplyMessageId : Uint4B
   LpcReplyMessageId As Long
   '+0x240 ReadClusterSize  : Uint4B
   ReadClusterSize As Long
   '+0x244 GrantedAccess    : Uint4B
   GrantedAccess As Long
   '+0x248 CrossThreadFlags : Uint4B
   '+0x248 Terminated       : Pos 0, 1 Bit
   '+0x248 DeadThread       : Pos 1, 1 Bit
   '+0x248 HideFromDebugger : Pos 2, 1 Bit
   '+0x248 ActiveImpersonationInfo : Pos 3, 1 Bit
   '+0x248 SystemThread     : Pos 4, 1 Bit
   '+0x248 HardErrorsAreDisabled : Pos 5, 1 Bit
   '+0x248 BreakOnTermination : Pos 6, 1 Bit
   '+0x248 SkipCreationMsg  : Pos 7, 1 Bit
   '+0x248 SkipTerminationMsg : Pos 8, 1 Bit
   CrossThreadFlags As Long
   '+0x24c SameThreadPassiveFlags : Uint4B
   '+0x24c ActiveExWorker   : Pos 0, 1 Bit
   '+0x24c ExWorkerCanWaitUser : Pos 1, 1 Bit
   '+0x24c MemoryMaker      : Pos 2, 1 Bit
   SameThreadPassiveFlags As Long
   '+0x250 SameThreadApcFlags : Uint4B
   '+0x250 LpcReceivedMsgIdValid : Pos 0, 1 Bit
   '+0x250 LpcExitThreadCalled : Pos 1, 1 Bit
   '+0x250 AddressSpaceOwner : Pos 2, 1 Bit
   SameThreadApcFlags As Long
   '+0x254 ForwardClusterOnly : UChar
   ForwardClusterOnly As Byte
   '+0x255 DisablePageFaultClustering : UChar
    DisablePageFaultClustering As Byte
End Type
