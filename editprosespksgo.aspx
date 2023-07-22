<%@ Page Language="VB" %>
<%@ Import Namespace="System" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data.OleDb" %> 
<%@ Import Namespace="System.Configuration" %>
<%@ Import Namespace="System.Web.Configuration" %>

<html>
<head>
<style>
.verticalhorizontal {
  margin: 0;
  position: absolute;
  top: 50%;
  left: 50%;
  transform: translate(-50%, -50%);
}
</style>
</head>
<body>
<%
    'Inherits System.Windows.Forms.Form
    'Create ADO.NET objects.
    Dim myConn As SqlConnection
    Dim myCmd As SqlCommand
	Dim myCmdi As SqlCommand
    Dim myReader As SqlDataReader
    'results As String
	Dim a, aaa, jmla, jmldata as integer

	'Dim pageno as Integer = 1
	Dim pageno as String
	Dim jmlpage as Integer = 1
	Dim pagerows as Integer = 20
	Dim rowsperpage as Integer = 20
	Dim hal as Integer = 1
	Dim b as Single = 0
	Dim c as Integer = 0
	Dim idx as Integer = 0
	Dim itop as Integer = 1
	Dim ii as Integer = 0
	Dim k as Integer
	Dim userid as Integer = 0
	Dim namauserx as String = ""

	Dim SQL0 As String
	Dim SQL1 As String
	Dim SQL2 As String
	Dim SQLjo As String
	
	Dim statuskc as Integer = 0
	Dim statuskcp as Integer = 0
	
	Dim user0 as String
	Dim jenis0 as String
	Dim perihal0 as String
	Dim pengirim0 as String
	Dim tujuan0 as String
	Dim tembusan0 as String
	Dim vendorx as String
	Dim jenisund0 as String
	Dim bspk0 as String
	Dim btunjuk0 as String
	Dim btahu0 as String
	Dim tgl0 as String
	Dim idq As String = ""
	Dim namauserq As String = ""
	Dim statusq As String = ""
	Dim jabatanq As String = ""
	Dim klpq As String = ""
	Dim unitq As String = ""
	Dim posisiq As String = ""
	Dim notinq As String = ""
	Dim sekreq As String = ""
	Dim divq As String = ""
	Dim posisix as String = ""
	Dim levelaksesq as String = ""
	Dim mohonid as Long = 0
	Dim adaid as Long = 0
	Dim nkpid as Long = 0
	Dim tglku as String = ""
	Dim jenisver as String
	Dim status as String
	Dim dataexist as Integer = 0
	Dim tglx as String
	Dim nomorq as String = ""
	Dim statusup as Integer = 0
	Dim judulada as String = ""
	Dim jenisada as String = ""
	Dim jenisang as String = ""
	Dim nominal as Long = 0
	Dim ket as String = ""
	Dim tahun, bulan, tglhari as Integer
	Dim prenamafile as String = ""
	Dim folder1, folder2 as String
	Dim file1, file2, file1z, file2z as String
	Dim namavendor as String = ""
	Dim statusnextx as Integer = 0
	Dim namafileabc as String = ""
	Dim namafilexx as String = ""
	Dim folderstr as String = ""

	user0 = Session("username")
	idq = Session("userid")
	namauserq = Session("namauser")
	jabatanq = Session("jabatan")
	klpq = Session("klp")
	unitq = Session("unit")
	posisiq = Session("posisi")
	posisix = Session("posisistr")
	notinq = Session("notin")
	divq = Session("divisi")

'Waktu -> prefix namafile
	'Dim s As String = DateTime.Now.ToString("yyyyMMdd")
	's = s & Hour(Now) & Minute(Now) & Second(Now)
	'namatemp = nama0.ToLower()
	'namatemp = namatemp.Replace(" ", "")
	'prenamafile = namatemp.Substring(0, 6) & s

	If NOT Request.Form("userid") is Nothing Then
		userid = Val(Request.Form("userid"))
	End If

	If NOT Request.Form("padaid") is Nothing Then
		adaid = Val(Request.Form("padaid"))
	End If
	
	If NOT Request.Form("spkid") is Nothing Then
		jenisund0 = Val(Request.Form("sanggahid"))
	End If
	
	If NOT Request.Form("edittgl") is Nothing Then
		tglku = Request.Form("edittgl")
		Dim tglstr() = tglku.Split("-")
		tglx = tglstr(2) & "-" & tglstr(1) & "-" & tglstr(0)
		if tglku.length() > 5 Then
			Dim tglarr(3) as String 
			tglarr = tglku.Split("-")
			tglku = tglarr(2) & "-" & tglarr(1) & "-" & tglarr(0)
			tahun = Cint(tglarr(2))
			bulan = Cint(tglarr(1))
			tglhari = Cint(tglarr(0))
		End If
	End If

	
	If NOT Request.Form("editketerangan") is Nothing Then
		ket = Request.Form("editketerangan")
		'ket = status.ToUpper()
	End If

	
	If NOT Request.Form("editjenis") is nothing Then
		jenis0 = Request.Form("editjenis")
	End If
	
	If NOT Request.Form("editperihal") is nothing Then
		perihal0 = Request.Form("editperihal")
	End If
	
	If NOT Request.Form("editpengirim") is nothing Then
		pengirim0 = Request.Form("editpengirim")
	End If

	If NOT Request.Form("edittujuan") is nothing Then
		tujuan0 = Request.Form("edittujuan")
	End If
	
	If NOT Request.Form("editvendor") is nothing Then
		vendorx = Request.Form("editvendor")
	End If


'response.write(statusnextx)
'-----------------------------------------------------------------------

	Dim fileold as String = ""
	If NOT Request.Form("oldnamafile") is nothing Then
		fileold = Request.Form("oldnamafile")
	End If
	
	If NOT Request.Form("edittgl") is nothing Then
		tgl0 = Request.Form("edittgl")
		if tgl0.length() > 5 Then
			Dim tglarr(3) as String 
			tglarr = tgl0.Split("-")
			tgl0 = tglarr(2) & "-" & tglarr(1) & "-" & tglarr(0)
			tahun = Cint(tglarr(2))
			bulan = Cint(tglarr(1))
		End If
	End If
	
	levelaksesq = Session("levelakses")
	'posisi20 = Left(posisix, 20)
	
'***************************************************************************************************************
'==================================================================================================================
' Pengadaan POC
'==================================================================================================================

	aaa = 1
    Try
        If aaa=1 Then
%>
<!--#include file="koneksi.aspx"-->
<%				
			
			myCmd = myConn.CreateCommand

'** Cek apakah data sudah ada
			SQLjo = "SELECT * FROM [pengadaan_pks] WHERE (pengadaanID=" & adaid & ")"
			myCmd.CommandText = SQLjo
			myConn.Open()
			myReader = myCmd.ExecuteReader()
            If myReader.HasRows Then
                While myReader.Read()
					dataexist = 1
				End While
			End If
			myConn.Close()

'==================================================================================================================
' File
'==================================================================================================================
'Waktu -> prefix namafile
			Dim s As String = DateTime.Now.ToString("yyyyMMdd")
			Dim tglinputq As String = DateTime.Now.ToString("yyyy-MM-dd")
			Dim jamsaiki As String = DateTime.Now.ToString("HH:mm:ss")
			Dim tahunstr as String = DateTime.Now.ToString("yyyy")
			Dim tahunq as Integer = Val(tahunstr)
			s = s & Hour(Now) & Minute(Now) & Second(Now)
			prenamafile = s

			folder1 = "./DOC/pengadaan/" & tahun & "/" & bulan & "/" & tglhari & "/"
			folder2 = "/DOC/pengadaan/" & tahun & "/" & bulan & "/" & tglhari & "/"
			
			Dim folderPath As String = Server.MapPath(folder1)
			 'Check whether Directory (Folder) exists.

			If(Not System.IO.Directory.Exists(folderPath)) Then
				'System.IO.Directory.CreateDirectory(folderPath)
				Directory.CreateDirectory(folderPath)
			End If

'========================================================================
' File - 1
'========================================================================
    'Save the File to the Directory (Folder).
			Dim exten As String = ""
			Dim postedFile As HttpPostedFile = Request.Files("edituploadImage1")
			
    'Check if File is available.
			Dim filearr(10) as String
			If postedFile IsNot Nothing And postedFile.ContentLength > 0 Then
				Dim namafilex = postedFile.FileName
				filearr = namafilex.Split(".")
				Dim filearrno as Integer = filearr.length
				Dim aa = filearrno - 2
				namafilexx = namafilex
				exten = System.IO.Path.GetExtension(postedFile.FileName)
				'namafilexx = prenamafile & namafilexx '& exten
				namafilexx = prenamafile & namafilexx.Replace(" ","")
				namafileabc = Server.MapPath(folder1) + namafilexx
				postedFile.SaveAs(namafileabc)
			End If
			file1 = namafilexx
			If file1.length()<1 Then
				file1z = ""
			ElseIf file1.length()>0 Then
				file1z = ",namafile1='" & file1 & "'"
			End If

'========================================================================
' File - 2
'========================================================================
    'Save the File to the Directory (Folder).
			postedFile = Request.Files("edituploadImage2")
			namafilexx = ""
    'Check if File is available.
			If postedFile IsNot Nothing And postedFile.ContentLength > 0 Then
				Dim namafilex = postedFile.FileName
				filearr = namafilex.Split(".")
				Dim filearrno as Integer = filearr.length
				Dim aa = filearrno - 2
				namafilexx = namafilex
				exten = System.IO.Path.GetExtension(postedFile.FileName)
				namafilexx = prenamafile & namafilexx.Replace(" ","")
				namafileabc = Server.MapPath(folder1) + namafilexx
				postedFile.SaveAs(namafileabc)
			End If
			file2 = namafilexx
			If file2.length()<1 Then
				file2z = ""
			ElseIf file2.length()>0 Then
				file2z = ",namafile2='" & file2 & "'"
			End If
			
			folderstr = ",namafolder='" & folder2 & "'"
'==================================================================================================================
' Masukkan ke database
'==================================================================================================================
			If dataexist=1 Then
				SQL0 = "UPDATE pengadaan_pks SET tgl='" & tglx & _
						"'" & folderstr & file1z & file2z & " WHERE (pengadaanID=" & adaid & ")"
'response.write(SQL0)
				myCmd.CommandText = SQL0
				myConn.Open()
				myReader = myCmd.ExecuteReader()
				myConn.Close()
			End If
			
        End If

    Catch ex As Exception
        'lbl1.Text="Error while connecting Server."
		'Response.Write (ex.Message)

    Finally
		'myReader.Close()
        'myConn.Close() 'Whether there is error or not. Close the connection.

    End Try
	
	Session("username") = user0
	Session("userid") = idq
	Session("namauser") = namauserq
	'Session("idemployee") = nppq
	Session("levelakses") = levelaksesq
	Session("klp") = klpq
	Session("unit") = unitq
	Session("posisi") = posisiq
	Session("posisistr") = posisix
	Session("divisi") = divq

			
%>
<img src="./images/spinnercolor.gif" alt="spinner" class="verticalhorizontal" height="200" width="200">
<meta http-equiv="refresh" content="0;url=prosespks.aspx?padaid=<% response.write(adaid) %>&userid=<% response.write(userid) %>&statusnext=<% response.write(statusnextx) %>" />
</body>
</html>


