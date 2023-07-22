<%@ Page Language="VB" %>
<%@ Import Namespace="System" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data.OleDb" %> 
<%@ Import Namespace="System.Configuration" %>
<%@ Import Namespace="System.Web.Configuration" %>

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
	Dim SQL3 As String
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
	Dim rubrikstr as String = ""
	Dim idku as Integer = 0
	
	Dim kelid0 as Integer = 0
	Dim rubrik0 as String = ""
	Dim divid0 as Integer = 1
	Dim suratmemo as String = ""

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



	If NOT Request.QueryString("kelid") is Nothing Then
		 kelid0 = Request.QueryString("kelid")
	End If
	
	If NOT Request.QueryString("rubrik") is Nothing Then
		 rubrik0 = Request.QueryString("rubrik")
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
			
			
'==============================================================================
' Cek apakah rubrik baru sudah ada
'==============================================================================
			SQL3 = "SELECT * FROM rubrik WHERE (rubrik='" & rubrik0 & "' AND kelompok=" & kelid0 & ")"
'response.write(SQL3 & "<br>")
			myCmd.CommandText = SQL3
			myConn.Open()
			myReader = myCmd.ExecuteReader()
			a = 0
			b = 0
			If myReader.HasRows Then
				While myReader.Read()
					a = 1
				End While
			End If
			myConn.Close()			
			a = 0
			
'==============================================================================
' Jika rubrik belum ada (a=0) -> BIKIN nomor
'==============================================================================
			If a = 0 Then
				rubrikstr = rubrik0 & "/"
				SQL0 = "INSERT INTO rubrik (rubrik, kelompok, status, divisi) VALUES('" & rubrikstr & "'," & kelid0 & ",1,1)"
'response.write(SQL0)			
				myCmd.CommandText = SQL0
				myConn.Open()
				myReader = myCmd.ExecuteReader()
				myConn.Close()
				b = 1
'=====================================================================================================================
' Update nomor pada pengadaan spk
'=====================================================================================================================				
			End If
							
			response.write(b)

        End If

    Catch ex As Exception
        'lbl1.Text="Error while connecting Server."
		'Response.Write (ex.Message)

    Finally
		'myReader.Close()
        'myConn.Close() 'Whether there is error or not. Close the connection.

    End Try
%>
