' This source-code is freely distributed as part of "DevCase for .NET Framework".
' 
' Maybe you would like to consider to buy this powerful set of libraries to support me. 
' You can do loads of things with my apis for a big amount of diverse thematics.
' 
' Here is a link to the purchase page:
' https://codecanyon.net/item/elektrokit-class-library-for-net/19260282
' 
' Thank you.

#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " HRESULT "

Namespace DevCase.Interop.Unmanaged.Win32.Enums

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Specifies error conditions and warning conditions about the success or failure of a method or function call.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <remarks>
    ''' Wikipedia: <see href="http://en.wikipedia.org/wiki/HRESULT"/> class.
    ''' <para></para>
    ''' MSDN guidelines: <see href="https://msdn.microsoft.com/en-us/library/windows/desktop/ff485842%28v=vs.85%29.aspx"/> class.
    ''' <para></para>
    ''' List of HRESULT codes: <see href="https://msdn.microsoft.com/en-us/library/cc704587.aspx"/> class.
    ''' </remarks>
    ''' ----------------------------------------------------------------------------------------------------
    Public Enum HResult As Integer

        ' *****************************************************************************
        '                            WARNING!, NEED TO KNOW...
        '
        '  THIS ENUMERATION IS PARTIALLY DEFINED TO MEET THE PURPOSES OF THIS API
        ' *****************************************************************************

        ''' <summary>
        ''' Success.
        ''' </summary>
        S_OK = &H0

        ''' <summary>
        ''' Success.
        ''' </summary>
        S_FALSE = &H1

        ''' <summary>
        ''' Access denied.
        ''' </summary>
        E_ACCESSDENIED = &H80070005

        ''' <summary>
        ''' Unspecified Error.
        ''' </summary>
        E_FAIL = &H80004005

        ''' <summary>
        ''' Invalid parameter value.
        ''' </summary>
        E_INVALIDARG = &H80070057

        ''' <summary>
        ''' 
        ''' </summary>
        E_ELEMENTNOTFOUND = &H80070490

        ''' <summary>
        ''' Out Of memory.
        ''' </summary>
        E_OUTOFMEMORY = &H8007000E

        ''' <summary>
        ''' NULL was passed incorrectly For a pointer value.
        ''' </summary>
        E_POINTER = &H80004003

        ''' <summary>
        ''' Unexpected condition.
        ''' </summary>
        E_UNEXPECTED = &H8000FFFF

        ''' <summary>
        ''' Operation aborted.
        ''' </summary>
        E_ABORT = &H80004004

        ''' <summary>
        ''' Handle that is not valid.
        ''' </summary>
        E_HANDLE = &H80070006

        ''' <summary>
        ''' Not implemented.
        ''' </summary>
        E_NOTIMPL = &H80004001

        '''<summary>
        '''The data necessary to complete this operation is not yet available.
        '''</summary>
        E_PENDING = &H8000000A

        ''' <summary>
        ''' 
        ''' </summary>
        DISP_E_OVERFLOW = &H8002000A

        ''' <summary>
        ''' A divide by zero error.
        ''' </summary>
        DISP_E_DIVBYZERO = &H80020012

        ''' <summary>
        ''' 
        ''' </summary>
        E_BOUNDS = &H8000000B

        ''' <summary>
        ''' 
        ''' </summary>
        E_CHANGED_STATE = &H8000000C

        ''' <summary>
        ''' 
        ''' </summary>
        ERROR_BAD_FORMAT = &H8007000B

        ''' <summary>
        ''' 
        ''' </summary>
        ERROR_ARITHMETIC_OVERFLOW = &H80070216

        ''' <summary>
        ''' 
        ''' </summary>
        ERROR_PATH_NOT_FOUND = &H80070003

        ''' <summary>
        ''' 
        ''' </summary>
        ERROR_FILE_NOT_FOUND = &H80070002

        ''' <summary>
        ''' 
        ''' </summary>
        ERROR_FILENAME_EXCED_RANGE = &H800700CE

        ''' <summary>
        ''' 
        ''' </summary>
        ERROR_STACK_OVERFLOW = &H800703E9

        ''' <summary>
        ''' The process cannot access the file because it is being used by another process.
        ''' </summary>
        ERROR_SHARING_VIOLATION = &H80070020

        ''' <summary>
        ''' 
        ''' </summary>
        FUSION_E_INVALID_NAME = &H80131047

        ''' <summary>
        ''' 
        ''' </summary>
        FUSION_E_REF_DEF_MISMATCH = &H80131040

        ''' <summary>
        ''' Error wrapped in <see cref="Global.System.Security.Cryptography.CryptographicException"/> class.
        ''' </summary>
        NTE_FAIL = &H80090020

        ''' <summary>
        ''' 
        ''' </summary>
        REGDB_E_CLASSNOTREG = &H80040154

        ''' <summary>
        ''' 
        ''' </summary>
        RO_E_CLOSED = &H80000013

        ''' <summary>
        ''' Error wrapped in <see cref="Global.System.InvalidCastException"/> class.
        ''' </summary>
        TYPE_E_TYPEMISMATCH = &H80028CA0

        ''' <summary>
        ''' 
        ''' </summary>
        TYPE_E_ELEMENTNOTFOUND = &H8002802B

        ''' <summary>
        ''' 
        ''' </summary>
        NO_OBJECT = &H800401E5

        ''' <summary>
        ''' 
        ''' </summary>
        ERROR_CANCELLED = &H800704C7

        ''' <summary>
        ''' 
        ''' </summary>
        COR_E_ABANDONEDMUTEX = &H8013152D

        ''' <summary>
        ''' 
        ''' </summary>
        COR_E_AMBIGUOUSMATCH = &H8000211D

        ''' <summary>
        ''' 
        ''' </summary>
        COR_E_APPDOMAINUNLOADED = &H80131014

        ''' <summary>
        ''' Error wrapped in <see cref="Global.System.ApplicationException"/> class.
        ''' </summary>
        COR_E_APPLICATION = &H80131600

        '''' <summary>
        '''' Error wrapped in <see cref="Global.System.ArgumentException"/> class.
        '''' </summary>
        'COR_E_ARGUMENT = HResult.E_INVALIDARG

        ''' <summary>
        ''' Error wrapped in <see cref="ArgumentOutOfRangeException"/> class.
        ''' </summary>
        COR_E_ARGUMENTOUTOFRANGE = &H80131502

        '''' <summary>
        '''' Error wrapped in <see cref="Global.System.ArithmeticException"/> class.
        '''' </summary>
        'COR_E_ARITHMETIC = HResult.ERROR_ARITHMETIC_OVERFLOW

        ''' <summary>
        ''' Error wrapped in <see cref="ArrayTypeMismatchException"/> class.
        ''' </summary>
        COR_E_ARRAYTYPEMISMATCH = &H80131503

        '''' <summary>
        '''' Error wrapped in <see cref="Global.System.BadImageFormatException"/> class.
        '''' </summary>
        'COR_E_BADIMAGEFORMAT = HResult.ERROR_BAD_FORMAT

        ''' <summary>
        ''' 
        ''' </summary>
        COR_E_CANNOTUNLOADAPPDOMAIN = &H80131015

        ''' <summary>
        ''' Error wrapped in <see cref="Global.System.ContextMarshalException"/> class.
        ''' </summary>
        COR_E_CONTEXTMARSHAL = &H80131504

        ''' <summary>
        ''' 
        ''' </summary>
        COR_E_CUSTOMATTRIBUTEFORMAT = &H80131605

        ''' <summary>
        ''' 
        ''' </summary>
        COR_E_DATAMISALIGNED = &H80131541

        ' ''' <summary>
        ' ''' Error wrapped in <see cref="Global.System.DivideByZeroException"/> class.
        ' ''' </summary>
        ' COR_E_DIVIDEBYZERO = HResult.DISP_E_DIVBYZERO

        ''' <summary>
        ''' 
        ''' </summary>
        COR_E_DLLNOTFOUND = &H80131524

        ''' <summary>
        ''' Error wrapped in <see cref="Global.System.DuplicateWaitObjectException"/> class.
        ''' </summary>
        COR_E_DUPLICATEWAITOBJECT = &H80131529

        ''' <summary>
        ''' Error wrapped in <see cref="Global.System.IO.EndOfStreamException"/> class.
        ''' </summary>
        COR_E_ENDOFSTREAM = &H80070026

        ''' <summary>
        ''' 
        ''' </summary>
        COR_E_FILELOAD = &H80131621

        '''' <summary>
        '''' Error wrapped in <see cref="Global.System.IO.FileNotFoundException"/> class.
        '''' </summary>
        'COR_E_FILENOTFOUND = HResult.ERROR_FILE_NOT_FOUND

        '''' <summary>
        '''' Error wrapped in <see cref="Global.System.IO.DirectoryNotFoundException"/> class.
        '''' </summary>
        'COR_E_DIRECTORYNOTFOUND = HResult.ERROR_PATH_NOT_FOUND

        '''' <summary>
        '''' Error wrapped in <see cref="Global.System.IO.PathTooLongException"/> class.
        '''' </summary>
        'COR_E_PATHTOOLONG = HResult.ERROR_FILENAME_EXCED_RANGE

        ''' <summary>
        ''' Error wrapped in <see cref="Global.System.Exception"/> class.
        ''' </summary>
        COR_E_EXCEPTION = &H80131500

        ''' <summary>
        ''' Error wrapped in <see cref="Global.System.ExecutionEngineException"/> class.
        ''' </summary>
        COR_E_EXECUTIONENGINE = &H80131506

        ''' <summary>
        ''' Error wrapped in <see cref="Global.System.FieldAccessException"/> class.
        ''' </summary>
        COR_E_FIELDACCESS = &H80131507

        ''' <summary>
        ''' Error wrapped in <see cref="Global.System.FormatException"/> class.
        ''' </summary>
        COR_E_FORMAT = &H80131537

        ''' <summary>
        ''' 
        ''' </summary>
        COR_E_HOSTPROTECTION = &H80131640

        ''' <summary>
        ''' Error wrapped in <see cref="Global.System.IndexOutOfRangeException"/> class.
        ''' </summary>
        COR_E_INDEXOUTOFRANGE = &H80131508

        ''' <summary>
        ''' 
        ''' </summary>
        COR_E_INSUFFICIENTEXECUTIONSTACK = &H80131578

        ''' <summary>
        ''' 
        ''' </summary>
        COR_E_INSUFFICIENTMEMORY = &H8013153D

        ''' <summary>
        ''' Error wrapped in <see cref="Global.System.InvalidCastException"/> class.
        ''' </summary>
        COR_E_INVALIDCAST = &H80004002

        ''' <summary>
        ''' Error wrapped in <see cref="Global.System.Runtime.InteropServices.InvalidComObjectException"/> class.
        ''' </summary>
        COR_E_INVALIDCOMOBJECT = &H80131527

        ''' <summary>
        ''' Error wrapped in <see cref="Global.System.Reflection.InvalidFilterCriteriaException"/> class.
        ''' </summary>
        COR_E_INVALIDFILTERCRITERIA = &H80131601

        ''' <summary>
        ''' Error wrapped in <see cref="Global.System.Runtime.InteropServices.InvalidOleVariantTypeException"/> class.
        ''' </summary>
        COR_E_INVALIDOLEVARIANTTYPE = &H80131531

        ''' <summary>
        ''' Error wrapped in <see cref="Global.System.InvalidOperationException"/> class.
        ''' </summary>
        COR_E_INVALIDOPERATION = &H80131509

        ''' <summary>
        ''' Error wrapped in <see cref="Global.System.IO.IOException"/> class.
        ''' </summary>
        COR_E_IO = &H80131620

        ''' <summary>
        ''' 
        ''' </summary>
        COR_E_INVALIDPROGRAM = &H8013153A

        ''' <summary>
        ''' 
        ''' </summary>
        COR_E_KEYNOTFOUND = &H80131577

        ''' <summary>
        ''' 
        ''' </summary>
        COR_E_MARSHALDIRECTIVE = &H80131535

        ''' <summary>
        ''' 
        ''' </summary>
        COR_E_MEMBERACCESS = &H8013151A

        ''' <summary>
        ''' Error wrapped in <see cref="Global.System.MethodAccessException"/> class.
        ''' </summary>
        COR_E_METHODACCESS = &H80131510

        ''' <summary>
        ''' Error wrapped in <see cref="Global.System.MissingFieldException"/> class.
        ''' </summary>
        COR_E_MISSINGFIELD = &H80131511

        ''' <summary>
        ''' Error wrapped in <see cref="Global.System.Resources.MissingManifestResourceException"/> class.
        ''' </summary>
        COR_E_MISSINGMANIFESTRESOURCE = &H80131532

        ''' <summary>
        ''' Error wrapped in <see cref="Global.System.MissingMemberException"/> class.
        ''' </summary>
        COR_E_MISSINGMEMBER = &H80131512

        ''' <summary>
        ''' Error wrapped in <see cref="Global.System.MissingMethodException"/> class.
        ''' </summary>
        COR_E_MISSINGMETHOD = &H80131513

        ''' <summary>
        ''' 
        ''' </summary>
        COR_E_MISSINGSATELLITEASSEMBLY = &H80131536

        ''' <summary>
        ''' Error wrapped in <see cref="Global.System.MulticastNotSupportedException"/> class.
        ''' </summary>
        COR_E_MULTICASTNOTSUPPORTED = &H80131514

        ''' <summary>
        ''' Error wrapped in <see cref="Global.System.NotFiniteNumberException"/> class.
        ''' </summary>
        COR_E_NOTFINITENUMBER = &H80131528

        ''' <summary>
        ''' Error wrapped in <see cref="Global.System.NotSupportedException"/> class.
        ''' </summary>
        COR_E_NOTSUPPORTED = &H80131515

        '''' <summary>
        '''' Error wrapped in <see cref="Global.System.NullReferenceException"/> class.
        '''' </summary>
        'COR_E_NULLREFERENCE = HResult.E_POINTER

        ''' <summary>
        ''' 
        ''' </summary>
        COR_E_OBJECTDISPOSED = &H80131622

        ''' <summary>
        ''' 
        ''' </summary>
        COR_E_OPERATIONCANCELED = &H8013153B

        '''' <summary>
        '''' Error wrapped in <see cref="Global.System.OutOfMemoryException"/> class.
        '''' </summary>
        'COR_E_OUTOFMEMORY = HResult.E_OUTOFMEMORY

        ''' <summary>
        ''' Error wrapped in <see cref="Global.System.OverflowException"/> class.
        ''' </summary>
        COR_E_OVERFLOW = &H80131516

        ''' <summary>
        ''' 
        ''' </summary>
        COR_E_PLATFORMNOTSUPPORTED = &H80131539

        ''' <summary>
        ''' Error wrapped in <see cref="Global.System.RankException"/> class.
        ''' </summary>
        COR_E_RANK = &H80131517

        ''' <summary>
        ''' Error wrapped in <see cref="Global.System.Reflection.ReflectionTypeLoadException"/> class.
        ''' </summary>
        COR_E_REFLECTIONTYPELOAD = &H80131602

        ''' <summary>
        ''' Error wrapped in <see cref="Global.System.Runtime.Remoting.RemotingException"/> class.
        ''' </summary>
        COR_E_REMOTING = &H8013150B

        ''' <summary>
        ''' Error wrapped in <see cref="Global.System.Reflection.ReflectionTypeLoadException"/> class.
        ''' </summary>
        COR_E_SERVER = &H8013150E

        ''' <summary>
        ''' 
        ''' </summary>
        COR_E_RUNTIMEWRAPPED = &H8013153E

        ''' <summary>
        ''' 
        ''' </summary>
        COR_E_SAFEARRAYRANKMISMATCH = &H80131538

        ''' <summary>
        ''' Error wrapped in <see cref="Global.System.Runtime.InteropServices.SafeArrayTypeMismatchException"/> class.
        ''' </summary>
        COR_E_SAFEARRAYTYPEMISMATCH = &H80131533

        ''' <summary>
        ''' 
        ''' </summary>
        COR_E_SAFEHANDLEMISSINGATTRIBUTE = &H80131623

        ''' <summary>
        ''' Error wrapped in <see cref="Global.System.Security.SecurityException"/> class.
        ''' </summary>
        COR_E_SECURITY = &H8013150A

        ''' <summary>
        ''' 
        ''' </summary>
        COR_E_SEMAPHOREFULL = &H8013152B

        ''' <summary>
        ''' Error wrapped in <see cref="Global.System.Runtime.Serialization.SerializationException"/> class.
        ''' </summary>
        COR_E_SERIALIZATION = &H8013150C

        '''' <summary>
        '''' Error wrapped in <see cref="Global.System.StackOverflowException"/> class.
        '''' </summary>
        'COR_E_STACKOVERFLOW = HResult.ERROR_STACK_OVERFLOW

        ''' <summary>
        ''' Error wrapped in <see cref="Global.System.Threading.SynchronizationLockException"/> class.
        ''' </summary>
        COR_E_SYNCHRONIZATIONLOCK = &H80131518

        ''' <summary>
        ''' Error wrapped in <see cref="Global.System.SystemException"/> class.
        ''' </summary>
        COR_E_SYSTEM = &H80131501

        ''' <summary>
        ''' Error wrapped in <see cref="Global.System.Reflection.TargetException"/> class.
        ''' </summary>
        COR_E_TARGET = &H80131603

        ''' <summary>
        ''' Error wrapped in <see cref="Global.System.Reflection.TargetInvocationException"/> class.
        ''' </summary>
        COR_E_TARGETINVOCATION = &H80131604

        ''' <summary>
        ''' Error wrapped in <see cref="Global.System.Reflection.TargetParameterCountException"/> class.
        ''' </summary>
        COR_E_TARGETPARAMCOUNT = &H8002000E

        ''' <summary>
        ''' Error wrapped in <see cref="Global.System.Threading.ThreadAbortException"/> class.
        ''' </summary>
        COR_E_THREADABORTED = &H80131530

        ''' <summary>
        ''' Error wrapped in <see cref="Global.System.Threading.ThreadInterruptedException"/> class.
        ''' </summary>
        COR_E_THREADINTERRUPTED = &H80131519

        ''' <summary>
        ''' 
        ''' </summary>
        COR_E_THREADSTART = &H80131525

        ''' <summary>
        ''' Error wrapped in <see cref="Global.System.Threading.ThreadStateException"/> class.
        ''' </summary>
        COR_E_THREADSTATE = &H80131520

        ''' <summary>
        ''' 
        ''' </summary>
        COR_E_THREADSTOP = &H80131521

        ''' <summary>
        ''' 
        ''' </summary>
        COR_E_TIMEOUT = &H80131505

        ''' <summary>
        ''' 
        ''' </summary>
        COR_E_TYPEACCESS = &H80131543

        ''' <summary>
        ''' Error wrapped in <see cref="Global.System.TypeInitializationException"/> class.
        ''' </summary>
        COR_E_TYPEINITIALIZATION = &H80131534

        ''' <summary>
        ''' Error wrapped in <see cref="Global.System.EntryPointNotFoundException"/> class 
        ''' and also <see cref="Global.System.TypeLoadException"/> class.
        ''' </summary>
        COR_E_TYPELOAD = &H80131522

        ''' <summary>
        ''' 
        ''' </summary>
        COR_E_TYPEUNLOADED = &H80131013

        '''' <summary>
        '''' 
        '''' </summary>
        'COR_E_UNAUTHORIZEDACCESS = HResult.E_ACCESSDENIED

        ''' <summary>
        ''' 
        ''' </summary>
        COR_E_UNSUPPORTEDFORMAT = &H80131523

        ''' <summary>
        ''' Error wrapped in <see cref="Global.System.Security.VerificationException"/> class.
        ''' </summary>
        COR_E_VERIFICATION = &H8013150D

        ''' <summary>
        ''' 
        ''' </summary>
        COR_E_WAITHANDLECANNOTBEOPENED = &H8013152C

        ''' <summary>
        ''' 
        ''' </summary>
        CORSEC_E_CRYPTO = &H80131430

        ''' <summary>
        ''' 
        ''' </summary>
        CORSEC_E_CRYPTO_UNEX_OPER = &H80131431

        ''' <summary>
        ''' 
        ''' </summary>
        CORSEC_E_MIN_GRANT_FAIL = &H80131417

        ''' <summary>
        ''' 
        ''' </summary>
        CORSEC_E_NO_EXEC_PERM = &H80131418

        ''' <summary>
        ''' 
        ''' </summary>
        CORSEC_E_POLICY_EXCEPTION = &H80131416

    End Enum

End Namespace

#End Region
