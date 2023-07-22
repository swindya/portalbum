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
	Dim nama0 as string = ""
	Dim posisi0 as String = ""
	Dim jenisund0 as String = ""
	Dim nominal0 as String = ""
	Dim rubrik0 as String = ""
	Dim nokredit0 as String = ""
	Dim nodebet0 as String = ""
	Dim tgl0 as String = ""
	Dim jenissurat0 as String = ""
	Dim tujuan0 as String = ""
	Dim namafile0 as String = ""
	Dim namafile1 as String = ""
	Dim namafile2 as String = ""
	Dim namafileold as String = ""

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
	Dim suratq As String = ""
	Dim sekreq As String = ""
	Dim divq As String = ""
	Dim posisix as String = ""
	Dim tahun as integer = 0
	Dim bulan as integer = 0
	Dim tglhari as Integer = 0
	Dim tahunold as integer = 0
	Dim bulanold as integer = 0
	Dim tglhariold as Integer = 0
	Dim filepath1 as String
	Dim filepath2 as String
	Dim folder1 as String
	Dim folder2 as String
	Dim filestr as String = ""
	Dim maxnu as String = ""
	Dim divstr as String = ""
	Dim tglold as String = ""
	Dim jmlnom as Long = 0
	Dim rvid as Long = 0

	user0 = Session("username")
	idq = Session("userid")
	namauserq = Session("namauser")
	jabatanq = Session("jabatan")
	klpq = Session("klp")
	unitq = Session("unit")
	posisiq = Session("posisi")
	posisix = Session("posisistr")
	suratq = Session("surat")
	divq = Session("divisi")

'Waktu -> prefix namafile
	'Dim s As String = DateTime.Now.ToString("yyyyMMdd")
	's = s & Hour(Now) & Minute(Now) & Second(Now)
	'namatemp = nama0.ToLower()
	'namatemp = namatemp.Replace(" ", "")
	'prenamafile = namatemp.Substring(0, 6) & s

	If NOT Request.Form("editrvid") is nothing Then
		rvid = Request.Form("editrvid")
	End If

	If NOT Request.Form("editjenis") is nothing Then
		jenis0 = Request.Form("editjenis")
	End If
	
	If NOT Request.Form("editrubrik") is nothing Then
		rubrik0 = Request.Form("editrubrik")
	End If
	
	If NOT Request.Form("editperihal") is nothing Then
		perihal0 = Request.Form("editperihal")
	End If
	
	If NOT Request.Form("editnorekkredit") is nothing Then
		nokredit0 = Request.Form("editnorekkredit")
	End If

	If NOT Request.Form("editnorekdebet") is nothing Then
		nodebet0 = Request.Form("editnorekdebet")
	End If
	
	If NOT Request.Form("oldtgl") is nothing Then
		tglold = Request.Form("oldtgl")
		if tglold.length() > 5 Then
			Dim tglarro(3) as String 
			tglarro = tglold.Split("-")
			tglold = tglarro(2) & "-" & tglarro(1) & "-" & tglarro(0)
			tahunold = Cint(tglarro(2))
			bulanold = Cint(tglarro(1))
			tglhariold = Cint(tglarro(0))
		End If
	End If
	
	If NOT Request.Form("oldnamafile") is nothing Then
		namafileold = Request.Form("oldnamafile")
	End If
	
	If NOT Request.Form("editnominal") is nothing Then
		nominal0 = Request.Form("editnominal")
		nominal0 = nominal0.Replace(".","")
		nominal0 = nominal0.Replace(",","")
		jmlnom = Val(nominal0)
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

	folder1 = "./DOC/register/" & tahun & "/" & bulan & "/" & tglhari & "/"
	
	Dim folderPath As String = Server.MapPath(folder1)
     'Check whether Directory (Folder) exists.

	If(Not System.IO.Directory.Exists(folderPath)) Then
		'System.IO.Directory.CreateDirectory(folderPath)
		Directory.CreateDirectory(folderPath)
	End If

    'Save the File to the Directory (Folder).
	Dim exten As String = ""
	Dim postedFile As HttpPostedFile = Request.Files("uploadImage1")
	Dim namafolderx as String = "/DOC/register/" & tahun & "/" & bulan & "/" & tglhari & "/"
	
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
			exten = System.IO.Path.GetExtension(postedFile.FileName)
			namafilexx = prenamafile & namafilexx & exten
			namafile0 = Server.MapPath(folder1) + namafilexx
'Check if file exists
			If System.IO.File.Exists(namafile0) Then
				hal=100
			Else
				postedFile.SaveAs(namafile0)
			End If
			
			filestr = ",namafile = '" & namafilexx & "', namafolder='" & namafolderx & "'"
'If no new file
		Else
			If tgl0 = tglold Then
				tahunold = tahun
				bulanold = bulan
				tglhariold = tglhari
'If tgl update new
			Else
				folder1 = "./DOC/register/" & tahunold & "/" & bulanold & "/" & tglhariold & "/"
				folder2 = "./DOC/register/" & tahun & "/" & bulan & "/" & tglhari & "/"
				namafile1 = Server.MapPath(folder1) + namafileold
				namafile2 = Server.MapPath(folder2) + namafileold
				My.Computer.FileSystem.MoveFile(namafile1,namafile2,FileIO.UIOption.AllDialogs,FileIO.UICancelOption.ThrowException)
				namafolderx = "/DOC/register/" & tahun & "/" & bulan & "/" & tglhari & "/"
				filestr = ", namafolder='" & namafolderx & "'"				
			End If			
		End If
    End If
	

'response.write(namafilexx)
	Dim divarr(10) as String
	divarr = rubrik0.Split("/")
	divstr = divarr(0)
	
'***************************************************************************************************************	

	aaa = 1
    Try
        If (aaa=1)
			'myConn = New SqlConnection("Data Source=DESKTOP-S7N8DF9;Initial Catalog=KCT;User ID=sa;Password=sa;")
%>
<!--#include file="koneksi.aspx"-->
<%				
			
			myCmd = myConn.CreateCommand
			maxu = 0
'** Cek max value of nomor urut
			'SQLjo = "SELECT MAX(nourut) AS maxnourut FROM [regvoucher] WHERE (rubrik='" & rubrik0 & "' AND tahun=" & tahun & ")"
			SQLjo = "SELECT * FROM [regvoucher] WHERE (ID=" & rvid & ")"

			myCmd.CommandText = SQLjo
			myConn.Open()
			myReader = myCmd.ExecuteReader()
            If myReader.HasRows Then
                While myReader.Read()
					maxnu = myReader("nourut").ToString()
					maxu = Val(maxnu)
				End While
			End If
			myConn.Close()
			
'=========================================================================================================
			Dim nomorq as String = rubrik0 & "/" & maxu
			SQL0 = "UPDATE regvoucher SET tgl='" & tgl0 & "', jenis='" & jenis0 & "', perihal='" & perihal0 & "', nominal=" & jmlnom & _
			filestr & ", norekkredit='" & nokredit0 & "', norekdebet='" & nodebet0 & "', nomor='" & nomorq & "', rubrik='" & _
			rubrik0 & "' WHERE (ID=" & rvid & ")"

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
<meta http-equiv="refresh" content="0;url=regvoucher.aspx" />
</body>
</html>