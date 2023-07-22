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
	Dim statusnext as Integer = 0
	
	Dim user0 as String
	Dim jenis0 as String
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
	Dim aanwijid as Integer = 0
	Dim tglku as String = ""
	Dim jenisver as String
	Dim status as String
	Dim dataexist as Integer = 0
	Dim tglx as String
	Dim statusup as Integer = 0
	Dim judulada as String = ""
	Dim jenisada as String = ""
	Dim jenisang as String = ""
	Dim nominal0 as String = ""
	Dim nominal as Long = 0
	Dim hps0 as String = ""
	Dim hps as Long = 0
	Dim ket as String = ""
	Dim tahun, bulan, tglhari as Integer
	Dim prenamafile as String = ""
	Dim folder1, folder2 as String
	Dim nfsurat, nfsuratz, nfip, nfipz, namafileabc as String
	Dim nf1, nf2, nf3, nf4, nf5, nf6, nf7, nf8 as String
	Dim nf1z, nf2z, nf3z, nf4z, nf5z, nf6z, nf7z, nf8z as String
	'Dim s as String = ""

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
	
	If NOT Request.Form("aanid") is Nothing Then
		aanwijid = Val(Request.Form("aanid"))
	End If
	
	If NOT Request.Form("edittgl") is Nothing Then
		tglku = Request.Form("edittgl")
		Dim tglstr1() = tglku.Split(" ")
		Dim tglstr() = tglstr1(0).Split("-")
		tglx = tglstr(2) & "-" & tglstr(1) & "-" & tglstr(0)
'response.write(tglx)
		tahun = Cint(tglstr(2))
		bulan = Cint(tglstr(1))
		tglhari = Cint(tglstr(0))
	End If

	If NOT Request.Form("editnama") is Nothing Then
		judulada = Request.Form("editnama")
	End If

	
	If NOT Request.Form("editnom") is Nothing Then
		nominal0 = Request.Form("editnom")
		nominal0 = nominal0.Replace(".","")
		nominal0 = nominal0.Replace(",","")
	End If
	
	If NOT Request.Form("edithps") is Nothing Then
		hps0 = Request.Form("edithps")
		hps0 = hps0.Replace(".","")
		hps0 = hps0.Replace(",","")
	End If

	'If NOT Request.Form("editjenisanggaran") is Nothing Then
	'	jenisang = Request.Form("editjenisanggaran")
	'End If
	
	'If NOT Request.Form("editjenis") is Nothing Then
	'	jenisada = Request.Form("editjenis")
	'End If
	
	If NOT Request.Form("editketerangan") is Nothing Then
		ket = Request.Form("editketerangan")
		'ket = status.ToUpper()
	End If
	
	If NOT Request.Form("usulnext") is nothing Then
		statusnext = Val(Request.Form("usulnext"))
	End If	
	
	levelaksesq = Session("levelakses")
	'posisi20 = Left(posisix, 20)
	
'***************************************************************************************************************
'==================================================================================================================
' Jenis verifikasi
'==================================================================================================================

	aaa = 1
    Try
        If aaa=1 Then
%>
<!--#include file="koneksi.aspx"-->
<%				
			
			myCmd = myConn.CreateCommand

'** Cek apakah data sudah ada
			SQLjo = "SELECT * FROM [pengadaan_prosesusulan] WHERE (pengadaanID=" & adaid & ")"
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
	Dim namafilexx As String = ""
	If postedFile IsNot Nothing And postedFile.ContentLength > 3 Then
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
	nf1 = namafilexx
	If nf1.length()<1 Then
		nf1z = ""
	ElseIf nf1.length()>0 Then
		nf1z = ",namafile1='" & nf1 & "'"
	End If

'========================================================================
' File - 2
'========================================================================
    'Save the File to the Directory (Folder).
	postedFile = Request.Files("edituploadImage2")
	namafilexx = ""
    'Check if File is available.
	If postedFile IsNot Nothing And postedFile.ContentLength > 3 Then
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

	nf2 = namafilexx
	If nf2.length()<1 Then
		nf2z = ""
	ElseIf nf2.length()>0 Then
		nf2z = ",namafile2='" & nf2 & "'"
	End If

'========================================================================
' File - 3
'========================================================================
    'Save the File to the Directory (Folder).
	postedFile = Request.Files("edituploadImage3")
	namafilexx = ""
    'Check if File is available.
	If postedFile IsNot Nothing And postedFile.ContentLength > 3 Then
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
	nf3 = namafilexx
	If nf3.length()<1 Then
		nf3z = ""
	ElseIf nf3.length()>0 Then
		nf3z = ",namafile3='" & nf3 & "'"
	End If
'========================================================================
' File - 4
'========================================================================
    'Save the File to the Directory (Folder).
	postedFile = Request.Files("edituploadImage4")
	namafilexx = ""
    'Check if File is available.
	If postedFile IsNot Nothing And postedFile.ContentLength > 3 Then
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
	nf4 = namafilexx
	If nf4.length()<1 Then
		nf4z = ""
	ElseIf nf4.length()>0 Then
		nf4z = ",namafile4='" & nf4 & "'"
	End If
'========================================================================
' File - 5
'========================================================================
    'Save the File to the Directory (Folder).
	postedFile = Request.Files("edituploadImage5")
	namafilexx = ""
    'Check if File is available.
	If postedFile IsNot Nothing And postedFile.ContentLength > 3 Then
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
	nf5 = namafilexx
	If nf5.length()<1 Then
		nf5z = ""
	ElseIf nf5.length()>0 Then
		nf5z = ",namafile5='" & nf5 & "'"
	End If
'========================================================================
' File - 6
'========================================================================
    'Save the File to the Directory (Folder).
	postedFile = Request.Files("edituploadImage6")
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
	nf6 = namafilexx
	If nf6.length()<1 Then
		nf6z = ""
	ElseIf nf6.length()>0 Then
		nf6z = ",namafile6='" & nf6 & "'"
	End If
'==================================================================================================================
' Masukkan ke database
'==================================================================================================================
	
			If dataexist=1 Then
				SQL0 = "UPDATE pengadaan_prosesusulan SET nilai=" & nominal0 & ", nilaihps=" & hps0 & ", tgl='" & tglx & _
				"', keterangan='" & ket & "', namafolder='" & _
				folder2 & "'" & nf1z & nf2z & nf3z & nf4z & nf5z & nf6z & " WHERE (pengadaanID=" & adaid & ")"
'response.write(SQL0)
				myCmd.CommandText = SQL0
				myConn.Open()
				myReader = myCmd.ExecuteReader()
				myConn.Close()
				
				SQL1 = "UPDATE pengadaan SET tglstep='" & tglx & "' WHERE (ID=" & adaid & ")"			
				myCmd.CommandText = SQL1
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
<meta http-equiv="refresh" content="0;url=prosesusulan.aspx?statusnext=<% response.write(statusnext) %>&padaid=<% response.write(adaid) %>" />
</body>
</html>