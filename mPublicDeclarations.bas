Attribute VB_Name = "mPublicDeclarations"
Option Explicit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This module lists all public variables and constants used in the workbook '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'TimeFrame helps me keep track of the "resolution" at which the GT data are requested and extracted
Public Enum TimeFrame
    TimeFrameDay = 1
    TimeFrameWeek
    TimeFrameMonth
    TimeFrameYear
End Enum
'BuildSheets specifies whether samplings must be extracted to separate worksheets, and on what basis
Public Enum BuildSheets
    BuildSheetsNone = 0
    BuildSheetsByQueryTerm
    BuildSheetsByRegion
    BuildSheetsByBoth
End Enum
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Toggle to prevent some settings from bothering me while I work on the code
Public Const bProductionVersion As Boolean = True
Public Const sProductionVersionString = "Public Const bProductionVersion As Boolean = True"
'Version information
Public Const iMainVersionNumber As Integer = 2
Public Const iSubVersionNumber As Integer = 2
Public Const iBuildVersionNumber As Integer = 0
Public Const sBuildString = "Public Const iBuildVersionNumber As Integer = "

'String components used to build the web query which will be sent to Google
'These two constants remain the same (constant!) regardless of interface (i.e., GTweb or GTe)
Public Const sReqStrTerms As String = "&terms="
Public Const sReqStrKey As String = "&key="

'Constants for the Google Trends Web interface
Public Const sReqStrGTW_URL_Base As String = "https://www.googleapis.com/trends/v1beta/"
Public Const sReqStrGTW_StartDate As String = "&restrictions.startDate="
Public Const sReqStrGTW_EndDate As String = "&restrictions.endDate="
'Public Const sReqStrGTEH_TimelReso As String = "&timelineResolution="
Public Const sReqStrGTW_GeoRestr As String = "&restrictions.geo="
Public Const sReqStrGTW_Cat As String = "&restrictions.category="
Public Const sReqStrGTW_Prop As String = "&restrictions.property="
Public Const sReqStrGTW_Fields As String = "&restrictions.fields="

'Constants for the Google Trends Extended for Health interface
Public Const sReqStrGTEH_URL_Base As String = "https://www.googleapis.com/trends/v1beta/timelinesForHealth?"
Public Const sReqStrGTEH_JSON_Request As String = "&alt=json"
Public Const sReqStrGTEH_StartDate As String = "&time.startDate="
Public Const sReqStrGTEH_EndDate As String = "&time.endDate="
Public Const sReqStrGTEH_TimelReso As String = "&timelineResolution="
Public Const sReqStrGTEH_GeoRestr As String = "&geoRestriction."
Public Const sReqStrGTEH_APIKey As String = "[API_Key]" 'This is not used in the actual URL, but replaces the key in the sample URL written to the log file

'Constant for the Account URL
Public Const sAccountURL = "https://console.developers.google.com/apis/api/trends.googleapis.com/overview?project="

'Other constants for string building
Public Const sQuote As String = """"  'Chr(34)   'Used to simplify creating formulas using quotes"
Public Const sColon As String = ":"   'Chr(58)   '":"
Public Const sLBrace As String = "{"  'Chr(123) '"{"
Public Const sRBrace As String = "}"  'Chr(125) '"}"
Public Const sBackslash As String = "\"

Public Const sSignaturePhrase As String = "This extraction was done using the Google Trends Information Extraction Tool " _
            & "created by Dr J Raubenheimer from the NHMRC-funded Translational Australian Clinical Toxicology Programme " _
            & "at the University of Sydney"

'xx
Public Const sWebLink As String = "https://github.com/TrueInsight/Google-Trends-Extraction-Tool"

'Used in mBlockPaste and elsewhere to tell users that pasting into the sheet has been blocked
Public Const sPasteDisallowMsg As String = "You cannot paste into this workbook (except in cell edit mode)"

'Variables (used to be hard coded constants) used to check that we do not exceed the daily quotas
'Public Const iMaxQueriesPD As Integer = 5000
'Public Const iMaxQPS As Integer = 2
'Public Const iMaxQueriesP100S As Integer = 200
Public iMaxQueriesPD As Integer
Public iMaxQPS As Integer
Public iMaxQueriesP100S As Integer

'bQuotaExceeded and bOtherHTTPError pick up errors in the sending of the HTTP request and the receiving of the response
' they are set in mRequestAndParseData and used in mGoogleTrendsInfoExtraction
Public bQuotaExceeded As Boolean    'bQuotaExceeded tests whether the quota for the day has been exceeded
Public bOtherHTTPError As Boolean   'bOtherHTTPError tests for other errors in sending the request or receiving its response
Public bInvalidArgumentError As Boolean     'Added 2020-09-30 {"error": {"code": 400,"message": "Request contains an invalid argument.","errors": [{"message": "Request contains an invalid argument.","domain": "global","reason": "badRequest"}],"status": "INVALID_ARGUMENT"}}

'Added for version 2
'Allows me to turn off the user reporting so that the tool can be automated externally (from a VBE module in another workbook)
' to complete multiple extractions (note which procedures in mButtonCode set this variable)
' Note that it does not turn of error reporting message boxes, so these can still interfere with automation
Public bShowCompletionMsgBoxes As Boolean

'Added version 2
'Stores the entry which will be written to the log when a GTe or GTw extraction is done
Public sLogEntry As String

Sub SetHardBounds()
    iMaxQueriesPD = fvReturnNameValue("MaxQueriesPerDay")
    iMaxQPS = fvReturnNameValue("MaxQueriesPerSecond")
    iMaxQueriesP100S = iMaxQPS * 100
End Sub
