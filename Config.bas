Attribute VB_Name = "Config"
Option Explicit

' Sheet names
Public Const AccountingPolicySummarySheetName As String = "APS"
Public Const TrialBalanceSheetName As String = "Trial Balance 1"

' File paths
Public Const PolicyWorkbookRelativePath As String = "\AccountingPolicy\accounting_policy.xlsx"

' Row and column indices
Public Const MainTopicRow As Integer = 5
Public Const startRow As Integer = 7

' Other constants
Public Const MinimumRowHeight As Integer = 15
