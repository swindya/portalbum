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
	Dim nama0 as string = ""
	Dim tahun0 as String = ""
	Dim posisi0 as String = ""
	Dim jenisanggaran0 as String = ""
	Dim jenisund0 as String = ""
	Dim nominal0 as String = ""
	Dim hps0 as String = ""
	Dim pengirim0 as String = ""
	Dim tgl0 as String = ""
	Dim jenisnotin0 as String = ""
	Dim tujuan0 as String = ""
	Dim rubrik0 as String = ""
	Dim namafileabc as String = ""

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
	
	Dim nfsurat as String = ""
	Dim nfip as String = ""
	Dim nftor as String = ""
	Dim nfrab as String = ""

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


	If NOT Request.Form("barujenis") is nothing Then
		jenis0 = Request.Form("barujenis")
	End If
	
	If NOT Request.Form("barunama") is nothing Then
		nama0 = Request.Form("barunama")
	End If
	
	If NOT Request.Form("barutahun") is nothing Then
		tahun0 = Request.Form("barutahun")
	End If
	
	If NOT Request.Form("barujenisanggaran") is nothing Then
		jenisanggaran0 = Request.Form("barujenisanggaran")
	End If
	
	If NOT Request.Form("barunom") is nothing Then
		nominal0 = Request.Form("barunom")
		nominal0 = nominal0.Replace(",","")
	End If
	
	If NOT Request.Form("baruhps") is nothing Then
		hps0 = Request.Form("baruhps")
		hps0 = hps0.Replace(",","")
	End If
	
	If NOT Request.Form("barutgl") is nothing Then
		tgl0 = Request.Form("barutgl")
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

	folder1 = "./DOC/pengadaan/" & tahun & "/" & bulan & "/" & tglhari & "/"
	folder2 = "/DOC/pengadaan/" & tahun & "/" & bulan & "/" & tglhari & "/"
	
	Dim folderPath As String = Server.MapPath(folder1)
     'Check whether Directory (Folder) exists.

	If(Not System.IO.Directory.Exists(folderPath)) Then
		'System.IO.Directory.CreateDirectory(folderPath)
		Directory.CreateDirectory(folderPath)
	End If

'========================================================================
' Surat Permohonan - 1
'========================================================================
    'Save the File to the Directory (Folder).
	Dim exten As String = ""
	Dim postedFile As HttpPostedFile = Request.Files("uploadImage1")
	'Dim namafolderx as String = "/DOC/pengadaan/" & tahun & "/" & bulan & "/"
	
    'Check if File is available.
	Dim filearr(10) as String
	Dim namafilexx As String = ""
	If postedFile IsNot Nothing And postedFile.ContentLength > 0 Then
		Dim namafilex = postedFile.FileName
		filearr = namafilex.Split(".")
		Dim filearrno as Integer = filearr.length
		Dim aa = filearrno - 2
		namafilexx = prenamafile & namafilex
		If aa > 0 Then
			namafilexx = ""
			For a=0 to aa
				namafilexx = namafilexx & filearr(a)
			Next a
			namafilexx = prenamafile & namafilexx & "." & filearr(aa+1)
		End If
		exten = System.IO.Path.GetExtension(postedFile.FileName)
		'namafilexx = prenamafile & namafilexx & exten
		namafileabc = Server.MapPath(folder1) + namafilexx
        postedFile.SaveAs(namafileabc)
    End If
	nfsurat = namafilexx

'========================================================================
' Ijin Prinsip - 2
'========================================================================
    'Save the File to the Directory (Folder).
	postedFile = Request.Files("uploadImage2")
	'Dim namafolderx as String = "/DOC/pengadaan/" & tahun & "/" & bulan & "/"
	namafilexx = ""
    'Check if File is available.
	If postedFile IsNot Nothing And postedFile.ContentLength > 0 Then
		Dim namafilex = postedFile.FileName
		filearr = namafilex.Split(".")
		Dim filearrno as Integer = filearr.length
		Dim aa = filearrno - 2
		namafilexx = prenamafile & namafilex
		If aa > 0 Then
			namafilexx = ""
			For a=0 to aa
				namafilexx = namafilexx & filearr(a)
			Next a
			namafilexx = prenamafile & namafilexx & "." & filearr(aa+1)
		End If
		exten = System.IO.Path.GetExtension(postedFile.FileName)
		'namafilexx = prenamafile & namafilexx & exten
		namafileabc = Server.MapPath(folder1) + namafilexx
        postedFile.SaveAs(namafileabc)
    End If
	nfip = namafilexx
'========================================================================
' TOR - 3
'========================================================================
    'Save the File to the Directory (Folder).
	postedFile = Request.Files("uploadImage3")
	'Dim namafolderx as String = "/DOC/pengadaan/" & tahun & "/" & bulan & "/"
	namafilexx = ""
    'Check if File is available.
	If postedFile IsNot Nothing And postedFile.ContentLength > 0 Then
		Dim namafilex = postedFile.FileName
		filearr = namafilex.Split(".")
		Dim filearrno as Integer = filearr.length
		Dim aa = filearrno - 2
		namafilexx = prenamafile & namafilex
		If aa > 0 Then
			namafilexx = ""
			For a=0 to aa
				namafilexx = namafilexx & filearr(a)
			Next a
			namafilexx = prenamafile & namafilexx & "." & filearr(aa+1)
		End If
		exten = System.IO.Path.GetExtension(postedFile.FileName)
		'namafilexx = prenamafile & namafilexx & exten
		namafileabc = Server.MapPath(folder1) + namafilexx
        postedFile.SaveAs(namafileabc)
    End If
	nftor = namafilexx
'========================================================================
' RAB - 4
'========================================================================
    'Save the File to the Directory (Folder).
	postedFile = Request.Files("uploadImage4")
	'Dim namafolderx as String = "/DOC/pengadaan/" & tahun & "/" & bulan & "/"
	namafilexx = ""
    'Check if File is available.
	If postedFile IsNot Nothing And postedFile.ContentLength > 0 Then
		Dim namafilex = postedFile.FileName
		filearr = namafilex.Split(".")
		Dim filearrno as Integer = filearr.length
		Dim aa = filearrno - 2
		namafilexx = prenamafile & namafilex
		If aa > 0 Then
			namafilexx = ""
			For a=0 to aa
				namafilexx = namafilexx & filearr(a)
			Next a
			namafilexx = prenamafile & namafilexx & "." & filearr(aa+1)
		End If
		exten = System.IO.Path.GetExtension(postedFile.FileName)
		'namafilexx = prenamafile & namafilexx & exten
		namafileabc = Server.MapPath(folder1) + namafilexx
        postedFile.SaveAs(namafileabc)
    End If
	nfrab = namafilexx
'========================================================================
	
'***************************************************************************************************************	

	aaa = 1
    Try
        If (aaa=1)
			'myConn = New SqlConnection("Data Source=DESKTOP-S7N8DF9;Initial Catalog=KCT;User ID=sa;Password=sa;")
%>
<!--#include file="koneksi.aspx"-->
<%				
			
			myCmd = myConn.CreateCommand

'** Cek max value of nomor urut
			SQLjo = "SELECT MAX(nourut) AS maxnourut FROM [notin] WHERE (rubrik='" & pengirim0 & "' AND tahun=" & tahun & ")"

			myCmd.CommandText = SQLjo
			myConn.Open()
			myReader = myCmd.ExecuteReader()
            If myReader.HasRows Then
                While myReader.Read()
					maxnu = myReader("maxnourut").ToString()
				End While
			End If
			myConn.Close()
			maxu = Val(maxnu) + 1
			
'=========================================
			SQL0 = "INSERT INTO pengadaan (tglmohon, klp, divisi, nama, jenispengadaan, jenisanggaran, nilai, nilaihps, tahun, " & _
					"namafolder, namafilesurat, namafileip, namafiletor, namafilerab, statuspass, nostep, step, createduser, " & _
					"createddate, statusrevisi, statusreject) VALUES('" & tgl0 & "'," & klpq & "," &	divq & ",'" & nama0 & "','" & jenis0 & _
					"','" & jenisanggaran0 & "'," & nominal0 & "," & hps0 & "," & tahun0 & ",'" & folder2 & "','" & nfsurat & "','" & nfip & _
					"','" & nftor & "','" & nfrab & "',0,1,'IJIN'," & idq & ",GETDATE(),0,0)"
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

<meta http-equiv="refresh" content="0;url=permohonan.aspx" />
</body>
</html>