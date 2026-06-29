
from enum import IntEnum


class ComError(IntEnum):
    """Common COM / HRESULT error codes returned by pywin32."""

    # --- Generic COM ---
    S_OK = 0
    E_FAIL = -2147467259                # 0x80004005
    E_INVALIDARG = -2147024809          # 0x80070057
    E_ACCESSDENIED = -2147024891        # 0x80070005
    E_NOTIMPL = -2147467263             # 0x80004001
    E_UNEXPECTED = -2147418113          # 0x8000FFFF
    E_OUTOFMEMORY = -2147024882         # 0x8007000E
    E_POINTER = -2147467261             # 0x80004003

    # --- Running Object Table / Moniker ---
    MK_E_UNAVAILABLE = -2147221021      # 0x800401E3
    MK_E_NOOBJECT = -2147221019         # 0x800401E5

    # --- Class registration / Dispatch ---
    CO_E_CLASSSTRING = -2147221005      # 0x800401F3
    REGDB_E_CLASSNOTREG = -2147221164   # 0x80040154
    CO_E_SERVER_EXEC_FAILURE = -2146959355  # 0x80080005

    # --- Outlook / Dispatch specific ---
    DISP_E_MEMBERNOTFOUND = -2147352573 # 0x80020003
    DISP_E_UNKNOWNNAME = -2147352570    # 0x80020006
    DISP_E_TYPEMISMATCH = -2147352571   # 0x80020005
    DISP_E_BADPARAMCOUNT = -2147352562  # 0x8002000E

    # --- RPC / Server died ---
    RPC_E_SERVERFAULT = -2147417851     # 0x80010105
    RPC_E_DISCONNECTED = -2147417848    # 0x80010108
    RPC_E_CALL_REJECTED = -2147418111   # 0x80010001

    @classmethod
    def from_code(cls, code: int):
        try:
            return cls(code)
        except ValueError:
            return None




class ExchangeStoreType(IntEnum):
    PRIMARY = 0
    DELEGATE = 1
    PUBLIC_FOLDER = 2
    ADDITIONAL = 3
    NOT_EXCHANGE = 4
    EXCHANGE_PUBLIC = 5
