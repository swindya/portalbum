<!DOCTYPE html>

<%@ Page Language="VB" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data.OleDb" %> 
<%@ Import Namespace="System.Configuration" %>
<%@ Import Namespace="System.Web.Configuration" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="ClosedXML.Excel" %>

<html>
<head>

<script language="vbscript" type="text/vbsscript" runat="server">

Function GetFullName$()
    Dim FirstName As String, LastName As String

    FirstName = "Sir"
    LastName = "Monay"

    Return LastName & ", " & FirstName
End Function

</script>

<title>Exercise</title>
</head>
<body>
<%
    Dim FullName$

    'FullName = GetFullName()
    'Response.Write(GetFullName$)
%>

<%
    'Using workBook As New XLWorkbook(FileUpload1.PostedFile.InputStream)

        'Read the first Sheet from Excel file.
		Dim exten As String = ""
		Dim namafile0 as String = ""
		Dim namafileout as String = ""
		Dim folder1 as String = "C:/Inetpub/wwwroot/portalbum/DOC/surat/Temp/"
		Dim folder2 as String = "./DOC/surat/Temp/"
		Dim postedFile As HttpPostedFile = Request.Files("uploadImage5")

    'Check if File is available.
		Dim filearr(10) as String
		Dim namafilexx As String = ""
		If postedFile IsNot Nothing Then
			If postedFile.ContentLength > 0 Then
				'Save the File.
				Dim namafilex = postedFile.FileName
				filearr = namafilex.Split(".")
				Dim filearrno as Integer = filearr.length
				Dim aa = filearrno - 2
				namafilexx = ""
				For a=0 to aa
					namafilexx = namafilexx & filearr(a)
				Next a
				'exten = System.IO.Path.GetExtension(uploadImage5.postedFile.FileName)
				namafilexx = namafilexx & ".xls"
				namafile0 = Server.MapPath(folder1) + namafilexx
				postedFile.SaveAs(namafile0)
			End If
		End If

namafile0 = folder1 + "testsurat.xls"
namafileout = folder1 + "testoutD.xls"
Dim strExcelConnr as String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & namafile0 & "';Extended Properties='Excel 12.0;HDR=Yes'"
Dim strExcelConnw as String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & namafileout & "';Extended Properties='Excel 12.0;HDR=Yes'"


Dim connExcel as new OleDbConnection(strExcelConnr)
Dim cmdExcel as new OleDbCommand()
cmdExcel.Connection = connExcel

Dim connExcelw as new OleDbConnection(strExcelConnw)
Dim cmdExcelw as new OleDbCommand()
cmdExcelw.Connection = connExcelw

'==============================================================================================

connExcel.Open()
Dim dtExcelSchema as New DataTable()
dtExcelSchema = connExcel.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Columns, Nothing)
connExcel.Close()

Dim i, j, jcol, jrow, krow, kcol, p, q, cola, rowa as Integer
Dim datax as String
Dim tglx(1000) as String
Dim jenisx(1000) as String
Dim perihalx(1000) as String
Dim rubrikx(1000) as String
Dim tujuanx(1000) as String
Dim tembusanx(1000) as String
Dim undanganx(1000) as String
Dim sqlout as String
Dim oda As New OleDbDataAdapter()
Dim ds as new DataSet()
Dim SheetName As String = dtExcelSchema.Rows(0)("TABLE_NAME").ToString()

connExcelw.Open()
cmdExcelw.CommandText = "CREATE TABLE Sheet1 (NOMOR Int,TGL Date,JENIS Int," & _
					"PERIHAL Text(200),RUBRIK Text(20),TUJUAN Text(200)," & _
					"TEMBUSAN Text(200), UNDANGAN Int)" 
cmdExcelw.ExecuteNonQuery()
connExcelw.Close()

connExcel.Open()
cmdExcel.CommandText = "SELECT * From [" + SheetName + "]"
oda.SelectCommand = cmdExcel
oda.Fill(ds)

jcol =  Val(ds.Tables(0).Columns.Count)
jrow =  Val(ds.Tables(0).Rows.Count)
krow = jrow - 1
kcol = jcol - 1
p = 0

For i = 0 To krow
	p = p + 1
	tglx(i) = ""
	jenisx(i)= ""
	perihalx(i) = ""
	rubrikx(i) = ""
	tujuanx(i) = ""
	tembusanx(i) = ""
	undanganx(i) = ""
	For j = 0 To kcol
	'datax = ds.Tables(0).Rows(i)(3).ToString().ToUpper()
		If Not ds.Tables(0).Rows(i)(j).ToString() is Nothing Then
			datax = ds.Tables(0).Rows(i)(j).ToString()
			If j = 1 Then
				if datax.length() > 5 Then
					Dim tgljarr(3) as String 
					tgljarr = datax.Split(" ")
					Dim tglarr(3) as String 
					tglarr = tgljarr(0).Split("/")
					tglx(i) = tglarr(2) & "-" & tglarr(1) & "-" & tglarr(0)
				End If
			Else if j = 2 Then
				jenisx(i) = ds.Tables(0).Rows(i)(j).ToString()
			Else if j = 3 Then
				perihalx(i) = ds.Tables(0).Rows(i)(j).ToString()
			Else if j = 4 Then
				rubrikx(i) = ds.Tables(0).Rows(i)(j).ToString()
			Else if j = 5 Then
				tujuanx(i) = ds.Tables(0).Rows(i)(j).ToString()
			Else if j = 6 Then
				tembusanx(i) = ds.Tables(0).Rows(i)(j).ToString()
			Else if j = 7 Then
				undanganx(i) = ds.Tables(0).Rows(i)(j).ToString()
			End If
			response.write(datax & " - ")
		End If
	Next j
		connExcelw.Open()
		cmdExcelw.Connection = connExcelw
		sqlout = "INSERT INTO [Sheet1$] (NOMOR,TGL,JENIS,PERIHAL,RUBRIK,TUJUAN,TEMBUSAN,UNDANGAN) " & _
				"VALUES(" & p & ",'" & tglx & "','" & jenisx & "','" & perihalx & "','" & _
				rubrikx & "','" & tujuanx & "','" & tembusanx & "','" & undanganx & "')"
response.write(sqlout)
		cmdExcelw.CommandText = sqlout
		cmdExcelw.ExecuteNonQuery()
		connExcelw.Close()
	response.write("<br>")
Next i
connExcel.Close()


namafileout = folder2 + "testoutD.xls"

%>
<a href="<% response.write(namafileout) %>" download rel="noopener noreferrer" target="_blank">
   Download File
</a>
</body>
</html>