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
	Dim tembusan0 as String = ""
	Dim pengirim0 as String = ""
	Dim tgl0 as String = ""
	Dim jenissurat0 as String = ""
	Dim tujuan0 as String = ""
	Dim rubrik0 as String = ""
	Dim namafile0 as String = ""
	Dim verid as Integer = 0
	Dim adaid as Integer = 0

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
	Dim regvid as String = ""

	Dim SQL0 As String
	Dim SQL1 As String
	Dim SQL2 As String
	Dim SQLdel As String
	
	Dim statuskc as Integer = 0
	Dim statuskcp as Integer = 0
	
	Dim regppid as String = ""
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
	
	Dim idxx as String = ""
	Dim nosuratxx as String = ""
	Dim tglxx as String = ""
	Dim tglinputxx as String = ""
	Dim nourutxx as String = ""
	Dim jenisxx as String = ""
	Dim klpxx as String = ""
	Dim divisixx as String = ""
	Dim rubrikxx as String = ""
	Dim perihalxx as String = ""
	Dim tujuanxx as String = ""
	Dim tahunxx as String = ""
	Dim tembusanxx as String = ""
	Dim jenisundxx as String = ""
	Dim cuserxx as String = ""
	Dim cdatexx as String = ""
	Dim namafilexx as String = ""
	Dim namafolderxx as String = ""
	Dim nominalxx as Long = 0
	
	Dim statusfile as Integer = 0

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

	If NOT Request.Form("verid") is nothing Then
		verid = Val(Request.Form("verid"))
	End If
	
	If NOT Request.Form("madaid") is nothing Then
		adaid = Val(Request.Form("madaid"))
	End If
	
	
'Waktu -> prefix namafile
	Dim s As String = DateTime.Now.ToString("yyyyMMdd")
	Dim tglsaiki As String = DateTime.Now.ToString("yyyy-MM-dd")
	Dim jamsaiki As String = DateTime.Now.ToString("HH:mm:ss")
	Dim tahunstr as String = DateTime.Now.ToString("yyyy")
	Dim tahunq as Integer = Val(tahunstr)
	s = s & Hour(Now) & Minute(Now) & Second(Now)
	prenamafile = s




'***************************************************************************************************************	

	aaa = 1
    Try
        If (aaa=1)
			'myConn = New SqlConnection("Data Source=DESKTOP-S7N8DF9;Initial Catalog=KCT;User ID=sa;Password=sa;")
%>
<!--#include file="koneksi.aspx"-->
<%				
			
			myCmd = myConn.CreateCommand
		
'=============================================================================================
' Baca data yg akan di hapus
'=============================================================================================
			SQL1 = "SELECT * FROM pengadaanverifikasi WHERE (ID=" & verid & ")"
'response.write(SQL1 & "<br>")					
			myCmd.CommandText = SQL1
			myConn.Open()
			myReader = myCmd.ExecuteReader()
			Dim tglsudate as Datetime
			a = 0
			If myReader.HasRows Then
				While myReader.Read()
					a = 1
				End While
			End If
			myConn.Close()

			If a=1 Then
'=============================================================================================
' Hapus data
'=============================================================================================				
				SQLdel = "DELETE FROM pengadaanverifikasi WHERE ID=" & verid
				myCmd.CommandText = SQLdel
				myConn.Open()
				myReader = myCmd.ExecuteReader()
				myConn.Close()
			End If

'response.write(SQLdel & "<br>")			
'=============================================================================================
			
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
<meta http-equiv="refresh" content="0;url=permohonanedit.aspx?padaid=<% response.write(adaid) %>&statusnext=1" />
</body>
</html>