Attribute VB_Name = "AssignCriticalitiesToTags"
'@Folder("VBAProject")

Option Explicit
Private tags As clsTags
Const wbCriticality As String = "WND Criticality Template.xlsx"
Private Disciplines As Collection

Sub LoadTags()

'Read in tags
' create workbook for each discipline

'foreach tag
    'lookup failure code output
    'set criticality using default MAH barrier, if defined
    'if isUtility then look at downgrade options / revising MAH barrier
    'if isSIL then set as LOPA/IPL in Non-fin business, which will give criticality A
    'if isSIS then set as LOPA/IPL in Non-fin business, which will give criticality A
    
    
    'copy results of template to row for tag into discipline workbook with comments
    '  / justification

'endforeach

End Sub


