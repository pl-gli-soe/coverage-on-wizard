VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InitConfigHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' FORREST SOFTWARE
' Copyright (c) 2015 Mateusz Forrest Milewski
'
' Permission is hereby granted, free of charge,
' to any person obtaining a copy of this software and associated documentation files (the "Software"),
' to deal in the Software without restriction, including without limitation the rights to
' use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software,
' and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
' INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
' IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
' WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
Public fup_code As String

Public coverageStockBasedOnQuestion As E_STOCK_FROM_WHERE

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

Public matchPUS As Boolean

Public do_we_want_to_run_coverage As Boolean
Public do_we_want_to_run_coord_list As Boolean
Public flats As Boolean


