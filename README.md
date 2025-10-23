# SAGE-Receipt-Upload
This macro will pick the data from Excel and enter in SAGE

Sub ReceiptUpload()

On Error GoTo ACCPACErrorHandler

Dim Session As New AccpacSession
Dim Signon As New AccpacSignonMgr
Dim ID As Long
Dim mDBLinkCmpRW As AccpacCOMAPI.AccpacDBLink
Session.Init "", "CS", "CS0001", "55A"

Session.Init "", "CS", "CS0001", "56A"

ID = Signon.Signon(Session)
Set mDBLinkCmpRW = Session.OpenDBLink(DBLINK_COMPANY, DBLINK_FLG_READWRITE)

Dim mDBLinkSysRW As AccpacCOMAPI.AccpacDBLink
Set mDBLinkSysRW = Session.OpenDBLink(DBLINK_SYSTEM, DBLINK_FLG_READWRITE)

Dim temp As Boolean
Dim ARRECEIPTS1batch As AccpacCOMAPI.AccpacView
Dim ARRECEIPTS1batchFields As AccpacCOMAPI.AccpacViewFields
mDBLinkCmpRW.OpenView "AR0041", ARRECEIPTS1batch
Set ARRECEIPTS1batchFields = ARRECEIPTS1batch.Fields

Dim ARRECEIPTS1header As AccpacCOMAPI.AccpacView
Dim ARRECEIPTS1headerFields As AccpacCOMAPI.AccpacViewFields
mDBLinkCmpRW.OpenView "AR0042", ARRECEIPTS1header
Set ARRECEIPTS1headerFields = ARRECEIPTS1header.Fields

Dim ARRECEIPTS1detail1 As AccpacCOMAPI.AccpacView
Dim ARRECEIPTS1detail1Fields As AccpacCOMAPI.AccpacViewFields
mDBLinkCmpRW.OpenView "AR0044", ARRECEIPTS1detail1
Set ARRECEIPTS1detail1Fields = ARRECEIPTS1detail1.Fields

Dim ARRECEIPTS1detail2 As AccpacCOMAPI.AccpacView
Dim ARRECEIPTS1detail2Fields As AccpacCOMAPI.AccpacViewFields
mDBLinkCmpRW.OpenView "AR0045", ARRECEIPTS1detail2
Set ARRECEIPTS1detail2Fields = ARRECEIPTS1detail2.Fields

Dim ARRECEIPTS1detail3 As AccpacCOMAPI.AccpacView
Dim ARRECEIPTS1detail3Fields As AccpacCOMAPI.AccpacViewFields
mDBLinkCmpRW.OpenView "AR0043", ARRECEIPTS1detail3
Set ARRECEIPTS1detail3Fields = ARRECEIPTS1detail3.Fields

Dim ARRECEIPTS1detail4 As AccpacCOMAPI.AccpacView
Dim ARRECEIPTS1detail4Fields As AccpacCOMAPI.AccpacViewFields
mDBLinkCmpRW.OpenView "AR0406", ARRECEIPTS1detail4
Set ARRECEIPTS1detail4Fields = ARRECEIPTS1detail4.Fields

Dim ARRECEIPTS1detail5 As AccpacCOMAPI.AccpacView
Dim ARRECEIPTS1detail5Fields As AccpacCOMAPI.AccpacViewFields
mDBLinkCmpRW.OpenView "AR0170", ARRECEIPTS1detail5
Set ARRECEIPTS1detail5Fields = ARRECEIPTS1detail5.Fields

ARRECEIPTS1batch.Compose Array(ARRECEIPTS1header)

ARRECEIPTS1header.Compose Array(ARRECEIPTS1batch, ARRECEIPTS1detail3, ARRECEIPTS1detail1, ARRECEIPTS1detail4, ARRECEIPTS1detail5)

ARRECEIPTS1detail1.Compose Array(ARRECEIPTS1header, ARRECEIPTS1detail2, Nothing)

ARRECEIPTS1detail2.Compose Array(ARRECEIPTS1detail1)

ARRECEIPTS1detail3.Compose Array(ARRECEIPTS1header)

ARRECEIPTS1detail4.Compose Array(ARRECEIPTS1header)

ARRECEIPTS1detail5.Compose Array(ARRECEIPTS1header)


Dim ARPAYMPOST2 As AccpacCOMAPI.AccpacView
Dim ARPAYMPOST2Fields As AccpacCOMAPI.AccpacViewFields
mDBLinkCmpRW.OpenView "AR0049", ARPAYMPOST2
Set ARPAYMPOST2Fields = ARPAYMPOST2.Fields


Dim ARRECMAC3batch As AccpacCOMAPI.AccpacView
Dim ARRECMAC3batchFields As AccpacCOMAPI.AccpacViewFields
mDBLinkCmpRW.OpenView "AR0041", ARRECMAC3batch
Set ARRECMAC3batchFields = ARRECMAC3batch.Fields

Dim ARRECMAC3header As AccpacCOMAPI.AccpacView
Dim ARRECMAC3headerFields As AccpacCOMAPI.AccpacViewFields
mDBLinkCmpRW.OpenView "AR0042", ARRECMAC3header
Set ARRECMAC3headerFields = ARRECMAC3header.Fields

Dim ARRECMAC3detail1 As AccpacCOMAPI.AccpacView
Dim ARRECMAC3detail1Fields As AccpacCOMAPI.AccpacViewFields
mDBLinkCmpRW.OpenView "AR0044", ARRECMAC3detail1
Set ARRECMAC3detail1Fields = ARRECMAC3detail1.Fields

Dim ARRECMAC3detail2 As AccpacCOMAPI.AccpacView
Dim ARRECMAC3detail2Fields As AccpacCOMAPI.AccpacViewFields
mDBLinkCmpRW.OpenView "AR0045", ARRECMAC3detail2
Set ARRECMAC3detail2Fields = ARRECMAC3detail2.Fields

Dim ARRECMAC3detail3 As AccpacCOMAPI.AccpacView
Dim ARRECMAC3detail3Fields As AccpacCOMAPI.AccpacViewFields
mDBLinkCmpRW.OpenView "AR0043", ARRECMAC3detail3
Set ARRECMAC3detail3Fields = ARRECMAC3detail3.Fields

Dim ARRECMAC3detail4 As AccpacCOMAPI.AccpacView
Dim ARRECMAC3detail4Fields As AccpacCOMAPI.AccpacViewFields
mDBLinkCmpRW.OpenView "AR0061", ARRECMAC3detail4
Set ARRECMAC3detail4Fields = ARRECMAC3detail4.Fields

Dim ARRECMAC3detail5 As AccpacCOMAPI.AccpacView
Dim ARRECMAC3detail5Fields As AccpacCOMAPI.AccpacViewFields
mDBLinkCmpRW.OpenView "AR0406", ARRECMAC3detail5
Set ARRECMAC3detail5Fields = ARRECMAC3detail5.Fields

Dim ARRECMAC3detail6 As AccpacCOMAPI.AccpacView
Dim ARRECMAC3detail6Fields As AccpacCOMAPI.AccpacViewFields
mDBLinkCmpRW.OpenView "AR0170", ARRECMAC3detail6
Set ARRECMAC3detail6Fields = ARRECMAC3detail6.Fields

ARRECMAC3batch.Compose Array(ARRECMAC3header)

ARRECMAC3header.Compose Array(ARRECMAC3batch, ARRECMAC3detail3, ARRECMAC3detail1, ARRECMAC3detail5, ARRECMAC3detail6)

ARRECMAC3detail1.Compose Array(ARRECMAC3header, ARRECMAC3detail2, ARRECMAC3detail4)

ARRECMAC3detail2.Compose Array(ARRECMAC3detail1)

ARRECMAC3detail3.Compose Array(ARRECMAC3header)

ARRECMAC3detail4.Compose Array(ARRECMAC3batch, ARRECMAC3header, ARRECMAC3detail3, ARRECMAC3detail1, ARRECMAC3detail2)

ARRECMAC3detail5.Compose Array(ARRECMAC3header)

ARRECMAC3detail6.Compose Array(ARRECMAC3header)


temp = ARRECMAC3batch.Exists
ARRECMAC3batch.RecordClear

ARRECMAC3batchFields("CODEPYMTYP").PutWithoutVerification ("CA")      ' Batch Type

ARRECMAC3headerFields("CODEPYMTYP").PutWithoutVerification ("CA")     ' Batch Type
ARRECMAC3detail3Fields("CODEPAYM").PutWithoutVerification ("CA")      ' Batch Type
ARRECMAC3detail1Fields("CODEPAYM").PutWithoutVerification ("CA")      ' Batch Type
ARRECMAC3detail2Fields("CODEPAYM").PutWithoutVerification ("CA")      ' Batch Type
ARRECMAC3detail4Fields("PAYMTYPE").PutWithoutVerification ("CA")      ' Batch Type
ARRECMAC3batchFields("CODEPYMTYP").PutWithoutVerification ("CA")      ' Batch Type

ARRECMAC3batchFields("CNTBTCH").PutWithoutVerification ("0")          ' Batch Number

temp = ARRECMAC3batch.Exists
ARRECMAC3batch.RecordCreate 1

ARRECMAC3batchFields("PROCESSCMD").PutWithoutVerification ("2")       ' Process Command

ARRECMAC3batch.Process
ARRECMAC3detail4.Cancel
temp = ARRECMAC3header.Exists
temp = ARRECMAC3header.Exists
ARRECMAC3header.RecordCreate 2
ARRECMAC3detail4.Cancel
temp = ARRECMAC3header.Exists
ARRECMAC3batch.Browse "((CODEPYMTYP = ""CA"") AND ((BATCHSTAT = 1) OR (BATCHSTAT = 7)))", 1

'Worksheets("sheet1").Activate
ARRECMAC3batchFields("BATCHDESC").PutWithoutVerification (Range("B1").Value)    ' Description

ARRECMAC3batch.Update

ARRECMAC3batchFields("DATEBTCH").Value = DateSerial(Year(Range("B2").Value), Month(Range("B2").Value), Day(Range("B2").Value))      ' Batch Date

ARRECMAC3batch.Update
ARRECMAC3headerFields("CNTITEM").PutWithoutVerification ("0")         ' Entry Number
temp = ARRECMAC3header.Exists
ARRECMAC3header.RecordCreate 2
ARRECMAC3detail4.Cancel
temp = ARRECMAC3header.Exists
ARRECMAC3detail4.Cancel
temp = ARRECMAC3header.Exists
ARRECMAC3batchFields("IDBANK").Value = (Range("B3").Value)            ' Bank Code
ARRECMAC3batch.Update
ARRECMAC3headerFields("CNTITEM").PutWithoutVerification ("0")         ' Entry Number
temp = ARRECMAC3header.Exists
ARRECMAC3header.RecordCreate 2
ARRECMAC3detail4.Cancel
temp = ARRECMAC3header.Exists

ARRECMAC3headerFields("RMITTYPE").Value = "2"                         ' Receipt Trans. Type
ARRECMAC3detail1.Cancel
ARRECMAC3headerFields("PROCESSCMD").PutWithoutVerification ("0")      ' Process Command Code
ARRECMAC3header.Process
temp = ARRECMAC3detail1.Exists
ARRECMAC3detail1.RecordClear
temp = ARRECMAC3detail1.Exists
ARRECMAC3detail1.RecordCreate 0
ARRECMAC3headerFields("RMITTYPE").Value = "1"                         ' Receipt Trans. Type
ARRECMAC3detail1.Cancel
ARRECMAC3headerFields("PROCESSCMD").PutWithoutVerification ("0")      ' Process Command Code
ARRECMAC3header.Process
ARRECMAC3detail4.Cancel

Range("A5").Activate
I = 5
Do While Len(Range("A" & I).Value) > 0

    ARRECMAC3headerFields("RMITTYPE").Value = "1"                         ' Receipt Trans. Type
    
    ARRECMAC3headerFields("IDCUST").Value = (Range("A" & I).Value)        ' Customer Number
    ARRECMAC3detail1.Cancel
'    ARRECMAC3headerFields("PROCESSCMD").PutWithoutVerification ("0")      ' Process Command Code
    ARRECMAC3header.Process
    ARRECMAC3detail4.Cancel
    ARRECMAC3headerFields("AMTRMIT").Value = (Range("B" & I).Value)       ' Bank Receipt Amount
    ARRECMAC3detail4Fields("PAYMTYPE").Value = "CA"                       ' Batch Type

'    ARRECMAC3detail4Fields("CNTBTCH").Value = "427"                      ' Batch Number
    ARRECMAC3detail4Fields("CNTITEM").Value = I - I                       ' Entry Number
    ARRECMAC3detail4Fields("IDCUST").Value = (Range("A" & I).Value)       ' ID Customer
    ARRECMAC3detail4Fields("AMTRMIT").Value = (Range("B" & I).Value)                   ' Receipt Amount
    ARRECMAC3detail4Fields("PROTYPE").PutWithoutVerification ("2")        ' Process Type
    ARRECMAC3headerFields("IDRMIT").Value = (Range("C" & I).Value)        ' Receipt Number
    ARRECMAC3headerFields("CODEPAYM").Value = (Range("D2").Value)         ' Payment Type
    
    
    ARRECMAC3detail4.Process
    ARRECMAC3header.Insert
    ARRECMAC3detail4.Cancel
    temp = ARRECMAC3header.Exists
    
    ARRECMAC3detail4Fields("PAYMTYPE").Value = "CA"                       ' Batch Type

    'ARRECMAC3detail4Fields("CNTBTCH").Value = "430"                       ' Batch Number
    ARRECMAC3detail4Fields("CNTITEM").Value = I - 4                         ' Entry Number
    ARRECMAC3detail4Fields("IDCUST").Value = (Range("A" & I).Value)        ' ID Customer
    ARRECMAC3detail4Fields("AMTRMIT").Value = (Range("B" & I).Value)                   ' Receipt Amount

    ARRECMAC3detail4.Process
    ARRECMAC3batch.Read
    ARRECMAC3headerFields("CNTITEM").PutWithoutVerification ("0")         ' Entry Number
    temp = ARRECMAC3header.Exists
    ARRECMAC3header.RecordCreate 2
    ARRECMAC3detail4.Cancel
    temp = ARRECMAC3header.Exists

    
    I = I + 1
Loop

Exit Sub

ACCPACErrorHandler:
  Dim lCount As Long
  Dim lIndex As Long

  If Errors Is Nothing Then
       MsgBox Err.Description
  Else
      lCount = Errors.Count

      If lCount = 0 Then
          MsgBox Err.Description
      Else
          For lIndex = 0 To lCount - 1
              MsgBox Errors.Item(lIndex)
          Next
          Errors.Clear
      End If
      Resume Next

  End If

End Sub

