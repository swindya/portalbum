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
	Dim nama0 as string = ""
	Dim divisi0 as string = ""
	Dim kel0 as Integer = 0
	Dim ada as String = ""
	Dim jenis0 as String = ""
	Dim vendor0 as String = ""
	Dim nominal0 as String = ""
	Dim per10 as String = ""
	Dim per20 as String = ""
	Dim tgl0 as String = ""
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
	Dim tglarr(3) as String

	Dim SQL0 As String
	Dim SQL1 As String
	Dim SQL2 As String
	Dim SQLjo As String
	
	Dim statuskc as Integer = 0
	Dim statuskcp as Integer = 0
	
	Dim user0 as String
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
	Dim jenisanggaran0 as Integer = 0
	Dim notinidx as Long = 0
	Dim divstr as String = ""
	Dim idku as Integer = 0
	Dim pksidx as Integer = 0
	Dim namadiv as String = ""
	Dim uic0 as Integer = 0

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

	If NOT Request.Form("barudivisi") is nothing Then
		divisi0 = Request.Form("barudivisi")
	End If
	
	If NOT Request.Form("barunama") is nothing Then
		nama0 = Request.Form("barunama")
	End If
	
	If NOT Request.Form("barujenis") is nothing Then
		jenis0 = Request.Form("barujenis")
	End If

	If NOT Request.Form("baruvendor") is nothing Then
		vendor0 = Request.Form("baruvendor")
	End If
	
	If NOT Request.Form("soflow13") is nothing Then
		uic0 = Val(Request.Form("soflow13"))
	End If
	
	If NOT Request.Form("barutgl1") is nothing Then
		per10 = Request.Form("barutgl1")
		if per10.length() > 5 Then 
			tglarr = per10.Split("-")
			per10 = tglarr(2) & "-" & tglarr(1) & "-" & tglarr(0)
		End If
	End If

	If NOT Request.Form("barutgl2") is nothing Then
		per20 = Request.Form("barutgl2")
		if per20.length() > 5 Then 
			tglarr = per20.Split("-")
			per20 = tglarr(2) & "-" & tglarr(1) & "-" & tglarr(0)
		End If
	End If
	
	If NOT Request.Form("barunominal") is nothing Then
		nominal0 = Request.Form("barunominal")
		nominal0 = nominal0.Replace(",","")
		nominal0 = nominal0.Replace(".","")
	End If
	
	If NOT Request.Form("barukel") is nothing Then
		kel0 = Val(Request.Form("barukel"))
	End If
	
	If NOT Request.Form("barutgl") is nothing Then
		tgl0 = Request.Form("barutgl")
		if tgl0.length() > 5 Then 
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

	'folder1 = "~/portalbum/DOC/surat/" & tahun & "/" & bulan & "/" & tglhari & "/"
	folder1 = "./DOC/surat/" & tahun & "/" & bulan & "/" & tglhari & "/"
	
	Dim folderPath As String = Server.MapPath(folder1)
     'Check whether Directory (Folder) exists.

	If(Not System.IO.Directory.Exists(folderPath)) Then
		'System.IO.Directory.CreateDirectory(folderPath)
		Directory.CreateDirectory(folderPath)
	End If

    'Save the File to the Directory (Folder).
	Dim exten As String = ""
	Dim postedFile As HttpPostedFile = Request.Files("uploadImage0")
	Dim namafolderx as String = "/DOC/surat/" & tahun & "/" & bulan & "/" & tglhari & "/"
	
    'Check if File is available.
	Dim filearr(10) as String
	Dim namafilexx As String = ""
	If postedFile IsNot Nothing And postedFile.ContentLength > 0 Then
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
    End If
	

'response.write(namafilexx)

'***************************************************************************************************************	

	aaa = 1
    Try
        If aaa=1 Then
			'myConn = New SqlConnection("Data Source=DESKTOP-S7N8DF9;Initial Catalog=KCT;User ID=sa;Password=sa;")
%>
<!--#include file="koneksi.aspx"-->
<%				
			
			myCmd = myConn.CreateCommand
			
'================================================================================
			SQL1 = "SELECT * FROM [divisi] WHERE div_id=" & divisi0			
			myCmd.CommandText = SQL1
			myConn.Open()
			myReader = myCmd.ExecuteReader()

			If myReader.HasRows Then
				While myReader.Read()
					namadiv = myReader("div_nama").ToString()
				End While
			End If
			myConn.Close()
'response.write(namadiv & "<br>")
'================================================================================
'** Cek max value of nomor urut
			SQLjo = "SELECT MAX(nourut) AS maxnourut FROM [surat_pks] WHERE (nomor LIKE '%" & namadiv & "%' AND tahun=" & tahun & ")"
			SQLjo = "SELECT MAX(nourut) AS maxnourut FROM [surat_pks] WHERE (nomor LIKE '%" & namadiv & "%' AND tahun=" & tahun & ")"

'response.write(SQLjo & "<br>")
			maxu = 0
			myCmd.CommandText = SQLjo
			myConn.Open()
			myReader = myCmd.ExecuteReader()
            If myReader.HasRows Then
                While myReader.Read()
					maxnu = myReader("maxnourut").ToString()
					maxu = Val(maxnu)
				End While
			End If
			myConn.Close()
			maxu = maxu + 1

'==================================================================================================================
' Masukkan ke surat_KP
'==================================================================================================================
			Dim nomorq as String = maxu & "/PKS/" & namadiv & "/" & tahun

			SQL0 = "INSERT INTO surat_pks (nourut, tglpks, nomor, tahun, klp, namapengadaan, vendor, jenisada, divisi, nominal," & _
					"per1, per2, jam, uic, createduser, createddate, namafile, namafolder) VALUES(" & _
					maxu & ",'" & tgl0 & "','" & nomorq & "'," & tahun & "," & kel0 & ",'" & nama0.ToUpper() & "','" & vendor0.ToUpper() & "','" & _
					jenis0 & "'," & divisi0 & "," & nominal0 & ",'" & per10 & "','" & per20 & "','" & jamsaiki & "'," & uic0 & "," & idq & ",'" & 
					tglinputq & "','" & namafilexx & "','" & namafolderx & "')"

'response.write(SQL0 & "<br>")			
			myCmd.CommandText = SQL0
			myConn.Open()
			myReader = myCmd.ExecuteReader()
			myConn.Close()

'==================================================================================================================
' Cari notinID -> Masukkan ke pengadaan
'==================================================================================================================
			SQL1 = "SELECT ID AS maxid FROM [surat_pks] WHERE (nomor='" & nomorq & "')"
			myCmd.CommandText = SQL1
			myConn.Open()
			myReader = myCmd.ExecuteReader()
            If myReader.HasRows Then
                While myReader.Read()
					pksidx = Val(myReader("maxid").ToString())
				End While
			End If
			myConn.Close()

'========================================================================================
		
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
<meta http-equiv="refresh" content="0;url=pksprint.aspx?id=<% response.write(pksidx) %>" />
</body>
</html>