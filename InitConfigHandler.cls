VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InitConfigHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public cbalFromWhere As E_CBAL_FROM_WHERE
Public pusFromWhere As E_PUS_FROM_WHERE


Public pusesForCoverage As E_TYPE_OF_PUSES_FOR_COVERAGE


'Public Enum E_CBAL_FROM_WHERE
'    E_CBAL_NA
'    E_CBAL_FROM_MGO
'    E_CBAL_FROM_W_GENERAL
'End Enum'
'
'
'Public Enum E_PUS_FROM_WHERE
'    E_PUS_NA
'    E_PUS_MGO
'    E_PUS_WIZARD
'    E_PUS_MIX
'End Enum
'
'Public Enum E_TYPE_OF_FLAT_TABLE
'    E_FLAT_NA
'    E_FLAT_RQM
'    E_FLAT_PUS
'    E_FLAT_RECV
'End Enum
'
'Public Enum E_TYPE_OF_RUN
'    E_TYPE_NA
'    E_FLATS
'    E_COV
'End Enum

Public rqmFlatTable As Boolean
Public pusFlatTable As Boolean
Public addRecv As Boolean

Public cbalFlatTable As Boolean


Public do_we_want_to_run_coverage As Boolean
Public do_we_want_to_run_coord_list As Boolean
Public flats As Boolean


