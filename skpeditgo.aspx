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
	Dim perihal0 as string = ""
	Dim ketuanama0 as string = ""
	Dim ketuajab0 as String = ""
	Dim wakanama0 as String = ""
	Dim wakajab0 as String = ""
	Dim anggota0 as String = ""
	Dim pengirim0 as String = ""
	Dim tgl0 as String = ""
	Dim jenisnotin0 as String = ""
	Dim tujuan0 as String = ""
	Dim rubrik0 as String = ""
	Dim namafile0 as String = ""

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
	Dim prenamafile as String = ""
	Dim namatemp as String = ""
	Dim jenisoutlet as String = ""
	Dim filePath as String = ""
	Dim kcx as String = ""
	Dim maxid as String = ""
	Dim maxu as Integer = 0

	Dim SQL0 As String
	Dim SQL1 As String
	Dim SQL2 As String
	Dim SQLjo As String
	
	Dim statuskc as Integer = 0
	Dim statuskcp as Integer = 0
	
	Dim skpid as String = ""
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
	Dim tahun as integer = 0
	Dim bulan as integer = 0
	Dim tglhari as Integer = 0
	Dim filepath1 as String
	Dim filepath2 as String
	Dim folder1 as String
	Dim folder2 as String
	Dim maxnu as String = ""
	
	Dim statusfile as Integer = 0

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

	If NOT Request.Form("editskpid") is nothing Then
		skpid = Request.Form("editskpid")
	End If
	
	If NOT Request.Form("editperihal") is nothing Then
		perihal0 = Request.Form("editperihal")
	End If

	If NOT Request.Form("editketua") is nothing Then
		ketuanama0 = Request.Form("editketua")
	End If
	
	If NOT Request.Form("editketuajab") is nothing Then
		ketuajab0 = Request.Form("editketuajab")
	End If

	If NOT Request.Form("editwaka") is nothing Then
		wakanama0 = Request.Form("editwaka")
	End If
	
	If NOT Request.Form("editwakajab") is nothing Then
		wakajab0 = Request.Form("editwakajab")
	End If

	If NOT Request.Form("editanggota") is nothing Then
		anggota0 = Request.Form("editanggota")
	End If
	
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
			tglhari = Cint(tglarr(0))
		End If
	End If
	
'Waktu -> prefix namafile
	Dim s As String = DateTime.Now.ToString("yyyyMMdd")
	Dim tglinputq As String = DateTime.Now.ToString("yyyy-MM-dd")
	Dim jamsaiki As String = DateTime.Now.ToString("HH:mm:ss")
	Dim tahunstr as String = DateTime.Now.ToString("yyyy")
	Dim tahunq as Integer = Val(tahunstr)
	s = s & Hour(Now) & Minute(Now) & Second(Now)
	prenamafile = s

	folder1 = "./DOC/notaintern/" & tahun & "/" & bulan & "/" & tglhari & "/"
	
	Dim folderPath As String = Server.MapPath(folder1)
     'Check whether Directory (Folder) exists.

	If(Not System.IO.Directory.Exists(folderPath)) Then
		'System.IO.Directory.CreateDirectory(folderPath)
		Directory.CreateDirectory(folderPath)
	End If

    'Save the File to the Directory (Folder).
	Dim exten As String = ""
	Dim postedFile As HttpPostedFile = Request.Files("uploadImage1")
	Dim namafolderx as String = "/DOC/notaintern/" & tahun & "/" & bulan & "/" & tglhari & "/"
	
    'Check if File is available.
	Dim filearr(10) as String
	Dim namafilexx As String = ""
	If postedFile IsNot Nothing Then
		If postedFile.ContentLength > 0 Then
			'Save the File.
			'Dim filePath1 As String = Server.MapPath("~/DOC/portalbum/") & Path.GetFileName(postedFile.FileName)
			'Dim filePath1 As String = Server.MapPath(folder1) + postedFile.FileName
			Dim namafilex = postedFile.FileName
			filearr = namafilex.Split(".")
			Dim filearrno as Integer = filearr.length
			Dim aa = filearrno - 2
			namafilexx = ""
			For a=0 to aa
				namafilexx = namafilexx & filearr(a)
'response.write(namafilexx & "<br>")
			Next a
			exten = System.IO.Path.GetExtension(postedFile.FileName)
			namafilexx = prenamafile & namafilexx & exten
			namafile0 = Server.MapPath(folder1) + namafilexx
			postedFile.SaveAs(namafile0)
			statusfile = 1
		End If
    End If
'***************************************************************************************************************	

	aaa = 1
    Try
        If (aaa=1)
			'myConn = New SqlConnection("Data Source=DESKTOP-S7N8DF9;Initial Catalog=KCT;User ID=sa;Password=sa;")
%>
<!--#include file="koneksi.aspx"-->
<%				
			
			myCmd = myConn.CreateCommand
		
'=========================================
			Dim nomorq as String = pengirim0 & maxu
			If statusfile=0 Then
				SQL0 = "UPDATE surat_kp SET ketua='" & ketuanama0.ToUpper() & "', ketua_jbtn='" & ketuajab0.ToUpper() & "', perihal='" & perihal0.ToUpper() & _
						"', wakilketua='" & wakanama0.ToUpper() & "', wakilketua_jbtn='" & wakajab0.ToUpper() & "', anggota='" & anggota0.ToUpper() & _
						"' WHERE (ID=" & skpid & ")"
			ElseIf statusfile=1 Then
				SQL0 = "UPDATE surat_kp SET ketua='" & ketuanama0.ToUpper() & "', ketua_jbtn='" & ketuajab0.ToUpper() & "', perihal='" & perihal0.ToUpper() & _
						"', wakilketua='" & wakanama0.ToUpper() & "', wakilketua_jbtn='" & wakajab0.ToUpper() & "', anggota='" & anggota0.ToUpper() & _
						"', namafile='" & namafilexx & "', namafolder='" & namafolderx & "' WHERE (ID=" & skpid & ")"
			End If

'kcx.ToUpper() & "',1)"
'response.write(SQL0)			
			myCmd.CommandText = SQL0
			myConn.Open()
			myReader = myCmd.ExecuteReader()
			myConn.Close()
			
        End If

    Catch ex As Exception
        'lbl1.Text="Error while connecting Server."
		'Response.Write (ex.Message)

    Finally
		'myReader.Close()
        'myConn.Close() 'Whether there is error or not. Close the connection.

    End Try
%>
<img src="./images/spinnercolor.gif" alt="spinner" class="verticalhorizontal" height="200" width="200">
<meta http-equiv="refresh" content="0;url=skp.aspx" />
</body>
</html>