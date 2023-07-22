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
	Dim perihal0 as string = ""
	Dim nama0 as string = ""
	Dim nomor0 as String = ""
	Dim dari0 as String = ""
	Dim posisi0 as String = ""
	Dim jenisund0 as String = ""
	Dim tembusan0 as String = ""
	Dim pengirim0 as String = ""
	Dim tgl0 as String = ""
	Dim tglsurat0 as String = ""
	Dim jenissurat0 as String = ""
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
	
	Dim tglarr(5) as String
	Dim tglarra(5) as String
	
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
	Dim filepath1 as String
	Dim filepath2 as String
	Dim folder1 as String
	Dim folder2 as String
	Dim maxnu as String = ""

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


	If NOT Request.Form("barujenis") is nothing Then
		jenis0 = Request.Form("barujenis")
	End If
	
	If NOT Request.Form("barunomor") is nothing Then
		nomor0 = Request.Form("barunomor")
	End If
	
	If NOT Request.Form("barudari") is nothing Then
		dari0 = Request.Form("barudari")
	End If
	
	If NOT Request.Form("baruperihal") is nothing Then
		perihal0 = Request.Form("baruperihal")
	End If
	
	If NOT Request.Form("barupengirim") is nothing Then
		pengirim0 = Request.Form("barupengirim")
	End If

	If NOT Request.Form("barutujuan") is nothing Then
		tujuan0 = Request.Form("barutujuan")
	End If
	
	If NOT Request.Form("barutembusan") is nothing Then
		tembusan0 = Request.Form("barutembusan")
	End If

	If NOT Request.Form("baruju") is nothing Then
		jenisund0 = Request.Form("baruju")
	End If
	
	If NOT Request.Form("barutgl") is nothing Then
		tgl0 = Request.Form("barutgl")
		if tgl0.length() > 5 Then
			tglarr = tgl0.Split("-")
			tgl0 = tglarr(2) & "-" & tglarr(1) & "-" & tglarr(0)
		End If
	End If
	
	If NOT Request.Form("barutglsurat") is nothing Then
		tglsurat0 = Request.Form("barutglsurat")
		if tglsurat0.length() > 5 Then
			tglarra = tglsurat0.Split("-")
			tglsurat0 = tglarra(2) & "-" & tglarra(1) & "-" & tglarra(0)
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

	'folder1 = "~/portalbum/DOC/surat/" & tahun & "/" & bulan & "/"
	folder1 = "./DOC/suratsekretaris/" & tahun & "/" & bulan & "/" & tglhari & "/"
	
	Dim folderPath As String = Server.MapPath(folder1)
     'Check whether Directory (Folder) exists.

	If(Not System.IO.Directory.Exists(folderPath)) Then
		'System.IO.Directory.CreateDirectory(folderPath)
		Directory.CreateDirectory(folderPath)
	End If

    'Save the File to the Directory (Folder).
	Dim exten As String = ""
	Dim postedFile As HttpPostedFile = Request.Files("uploadImage0")
	Dim namafolderx as String = "/DOC/suratsekretaris/" & tahun & "/" & bulan & "/" & tglhari & "/"
	
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
	response.write(namafilexx & "<br>")
			Next a
			exten = System.IO.Path.GetExtension(postedFile.FileName)
			namafilexx = prenamafile & namafilexx & exten
			namafile0 = Server.MapPath(folder1) + namafilexx
			postedFile.SaveAs(namafile0)
		End If
    End If
	

'response.write(namafilexx)
	
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
			'SQLjo = "SELECT MAX(nourut) AS maxnourut FROM [surat_sekretaris] WHERE (rubrik='" & pengirim0 & "' AND tahun=" & tahun & ")"

			'myCmd.CommandText = SQLjo
			'myConn.Open()
			'myReader = myCmd.ExecuteReader()
            'If myReader.HasRows Then
            '    While myReader.Read()
			'		maxnu = myReader("maxnourut").ToString()
			'	End While
			'End If
			'myConn.Close()
			'maxu = Val(maxnu) + 1
			
'=========================================
			Dim nomorq as String = pengirim0 & maxu
			SQL0 = "INSERT INTO surat_sekretaris (tglsurat, jenisstr, tahun, unit, klp, div, createduser, dari, tujuan," & _
					"perihal, cc, nomor, jam, createddate, namafile, namafolder) VALUES('" & tglsurat0 & "','" & _
					jenis0 & "'," & tahun & "," & unitq & "," & klpq & "," & divq & "," & idq & ",'" & dari0.ToUpper() & "','" & tujuan0.ToUpper() & _
					"','" & perihal0.ToUpper() & "','" & tembusan0.ToUpper() & "','" & nomor0 & "','" & jamsaiki & "','" & _
					tglinputq & "','" & namafilexx & "','" & namafolderx & "')"
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

<meta http-equiv="refresh" content="0;url=suratsekre.aspx?jenis=<% response.write(jenis0) %>" />
</body>
</html>