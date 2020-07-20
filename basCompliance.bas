Attribute VB_Name = "basCompliance"
Option Explicit
Dim regPattern As String
Dim code As String

Public Function GetFromOutlook()

Dim OutlookApp As Outlook.Application
Dim OutlookNamespace As Namespace
Dim olFolder As MAPIFolder
Dim OutlookMail As Variant
Dim objOwner As Object

Dim i As Long
Dim rCount As Long

Dim strPlanID As String
Dim strDepartment As String
Dim strInitials As String
Dim strManager As String
Dim strAssociateRole As String
Dim strSONI As String
Dim strSONIArticle As String


Set OutlookApp = New Outlook.Application
Set OutlookNamespace = OutlookApp.GetNamespace("MAPI")
Set objOwner = OutlookNamespace.CreateRecipient("RPS_Compliance_Referral@capgroup.com")

objOwner.Resolve

If objOwner.Resolved Then


Set olFolder = OutlookNamespace.GetSharedDefaultFolder(objOwner, olFolderInbox).Folders("TestStella")

Stop

End If

rCount = 2

For Each OutlookMail In olFolder.Items


    If OutlookMail.ReceivedTime >= Range("A1").Value Then
Dim Reg1 As New RegExp


' ## Case Statement to get tablular data
For i = 1 To 7

    Select Case i

    Case 1
    regPattern = "(PLAN ID(S)[:](.*))\n"

    Case 2
    regPattern = "(DEPARTMENT[:](.*))\n"

     Case 3
    regPattern = "(INITIALS[:](.*))\n"

     Case 4
    regPattern = "(MANAGER[:](.*))\n"

    Case 5
    regPattern = "(ASSOCIATE ROLE[:](.*))\n"

    Case 6
    regPattern = "(SONI[:](.*))\n"

    Case 7
    regPattern = "(SONI ARTICLE[:](.*))\n"



    End Select



    ExtractText (OutlookMail.Body)
Debug.Print "Code: " & code
    If i = 1 Then strPlanID = code
    If i = 2 Then strDepartment = code
    If i = 3 Then strInitials = code
    If i = 4 Then strManager = code
    If i = 5 Then strAssociateRole = code
    If i = 6 Then strSONI = code
    If i = 7 Then strSONIArticle = code

    Next i

'End tabular

    Range("A" & rCount) = OutlookMail.Subject
    Range("B" & rCount) = OutlookMail.ReceivedTime
    Range("C" & rCount) = strInitials
    Range("D" & rCount) = strManager
    Range("E" & rCount) = strAssociateRole
    Range("F" & rCount) = strDepartment
    Range("G" & rCount) = strPlanID
    Range("H" & rCount) = "Yes"
    Range("I" & rCount) = strSONI
    Range("J" & rCount) = strSONIArticle
    Range("K" & rCount) = OutlookMail.ReceivedTime




       rCount = rCount + 1
    End If
Next OutlookMail



Set olFolder = Nothing
Set OutlookNamespace = Nothing
Set OutlookApp = Nothing
'
End Function

Function ExtractText(Str As String) ' As String
 Dim regEx As New RegExp
 Dim NumMatches As MatchCollection
 Dim M As Match

 regEx.Pattern = regPattern

 Set NumMatches = regEx.Execute(Str)
 If NumMatches.Count = 0 Then
      ExtractText = ""
 Else
 Set M = NumMatches(0)
    ExtractText = M.SubMatches(1)
    Debug.Print "ExtractText: " & ExtractText
 End If
 code = ExtractText
 End Function
