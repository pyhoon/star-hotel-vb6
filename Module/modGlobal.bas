Attribute VB_Name = "modGlobalVariable"
' Version : 1.0.18
'
' Modified On : 19/05/2018
Option Explicit

Public CrApplication As CRAXDRT.Application

Public gstrCompanyID As String
Public gstrCompanyName As String
Public gstrCompanyName2 As String
Public gstrStreetAddress As String
Public gstrContactNo As String
Public gstrCurrency As String
Public gdtmFiscalYearStart As Date

Public gstrUserID As String
Public gstrUserName As String
Public gstrUserPassword As String
Public gstrUserSalt As String
Public gintUserIdle As Integer
Public gintUserGroup As Integer
Public gblnUserChangePassword As Boolean

Public glngReportNum As Long
Public gstrReportFileName As String
Public gstrReportTitle As String

Public gstrSQL As String

Public Const COMPANY_PRODUCT_NAME = "Hotel Booking System"
Public Const MOD_DASHBOARD = 1
Public Const MOD_BOOKING = 2
'Public Const MOD_BOOKING_VOID = 3
Public Const MOD_REPORT_LIST = 3
Public Const MOD_REPORT_PRINT = 4
Public Const MOD_REPORT_EXPORT = 5
Public Const MOD_REPORT_EDIT = 6
Public Const MOD_REPORT_EDIT_EXPERT = 7
Public Const MOD_FIND_CUSTOMER = 8
Public Const MOD_MAINTAIN_ROOM = 9
Public Const MOD_MAINTAIN_USER = 10
Public Const MOD_ACCESS_CONTROL = 11
Public Const REP_DAILY_BOOKING = 12
Public Const REP_WEEKLY_BOOKING = 13
Public Const REP_MONTHLY_BOOKING = 14
Public Const REP_WEEKLY_DEPOSIT_GRAPH = 15
Public Const REP_SHIFT_FOR_USER = 16
Public Const REP_SHIFT_ALL_USER = 17
