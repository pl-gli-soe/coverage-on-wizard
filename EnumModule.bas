Attribute VB_Name = "EnumModule"
Public Enum E_PUS_CZY_RQM_CZY_CBAL
    FOMULARZ_WYBORU_PLIKU_DLA_PUS
    FOMULARZ_WYBORU_PLIKU_DLA_RQM
    FOMULARZ_WYBORU_PLIKU_DLA_CBAL
End Enum


Public Enum E_CBAL_LIST
    CBAL_PLT = 1
    CBAL_PN
    CBAL_CBAL
End Enum

Public Enum E_RQMS_LIST
    RQMS_PLT = 1
    RQMS_PN
    RQMS_FUP_CODE
    RQMS_CW
    RQMS_QTY
End Enum


' ENUM FOR PUSes
' PLT PN  FUP_CODE    PUS_DATE    DEL_DATE    QTY DEL QTY RECV    BOOL RECV   PUS_NAME    DUNS    SUPPLIER NAME   ON MGO  ON WIZARD
' ==================================================================================================

Public Enum E_PUSES_LIST
    PUSES_PLT = 1
    PUSES_PN
    PUSES_FUP
    PUSES_PUS_DATE
    PUSES_DEL_DATE
    PUSES_QTY
    PUSES_DEL_QTY
    PUSES_RECV
    PUSES_BOOL_RECV
    PUSES_PUS_NAME
    PUSES_DUNS
    PUSES_SUPP_NM
    PUSES_ON_MGO
    PUSES_ON_WIZARD
    PUSES_LOG
End Enum
' ==================================================================================================

' recv type section
' ---------------------
Public Enum E_RECV_TYPE
    RECV_TBD = 0
    ON_ZERO = 1
    INLINE_WITH_QTY = 2
    NOT_INLINE_WITH_QTY = 3
    IN_TRANSIT = 4
End Enum
' ---------------------

Public Enum E_CBAL_FROM_WHERE
    E_CBAL_NA
    E_CBAL_FROM_MGO
    E_CBAL_FROM_W_GENERAL
    E_CBAL_FROM_WIZARD
End Enum


Public Enum E_PUS_FROM_WHERE
    E_PUS_NA
    E_PUS_MGO
    E_PUS_WIZARD
    E_PUS_MIX
End Enum

Public Enum E_TYPE_OF_PUSES_FOR_COVERAGE
    E_TYPE_PUS_WIZARD
    E_TYPE_PUS_MGO
End Enum

Public Enum E_TYPE_OF_FLAT_TABLE
    E_FLAT_NA
    E_FLAT_RQM
    E_FLAT_PUS
    E_FLAT_CBAL
End Enum

Public Enum E_TYPE_OF_RUN
    E_TYPE_NA
    E_FLATS
    E_COV
End Enum





' WIZARD SECTION
' ==================================================================================================
Public Enum E_PUS_SH
    O_INDX = 1
    O_PN
    O_DUNS
    O_FUP_code
    O_Pick_up_date
    O_Delivery_Date
    O_Pick_up_Qty
    O_PUS_Number
End Enum

Public Enum E_NEW_PROJECT_ITEM
    plt = 1
    PROJECT
    BIW_GA ' BIW or GA
    MY
    PHAZE
    BOM
    PICKUP_DATE
    PPAP_GATE
    mrd
    BUILD_START
    BUILD_END
    KOORDYNATOR
    E_ACTIVE
    CAPACITY_CHECK
    E_MRD_DATE
    E_MRD_REG_ROUTES
    E_PLATFORM
    E_TRANSPORTATION_ACCOUNT_NUMBER
    E_UNIQUE_ID
End Enum



Public Enum E_MASTER_MANDATORY_COLUMNS
    pn = 1
    Alternative_PN
    PN_Name
    GPDS_PN_Name
    duns
    Supplier_Name
    COUNTRY_CODE
    MGO_code
    Responsibility
    fup_code
    SQ
    ppap_status
    SQ_Comments
    MRD1_QTY
    MRD2_QTY
    Total_QTY
    ADD_to_T_slash_D
    MRD1_Ordered_date
    MRD1_Ordered_QTY
    MRD1_Ordered_STATUS
    MRD1_confirmed_qty
    MRD1_confirmed_qty_dot__Status
    MRD1_Total_PUS_STATUS
    MRD2_Ordered_date
    MRD2_Ordered_QTY
    MRD2_Ordered_STATUS
    MRD2_confirmed_qty
    MRD2_confirmed_qty_dot__Status
    MRD2_Total_PUS_STATUS
    Delivery_confirmation
    First_Confirmed_PUS_Date
    Delivery_reconfirmation
    Total_PUS_QTY
    Total_PUS_STATUS
    Comments
    Bottleneck
    Future_Osea
    DRE
    EDI_Received
    Capacity
    Oncost_confirmation
    BLANK3
    BLANK4
End Enum


