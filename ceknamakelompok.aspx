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
	Dim namadiv as string = ""
	Dim kodediv as string = ""
	Dim div1 as String = ""
	Dim kel1 as String = ""
	Dim unit1 as String = ""
	Dim posisi1 as String = ""
	Dim level1 as String = ""
	Dim csurat1 as String = ""
	Dim csekre1 as String = ""
	Dim cada1 as String = ""
	Dim caset1 as String = ""
	Dim pwd11 as String = ""
	Dim pwd21 as String = ""

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
	Dim divid as Integer = 0
	Dim namaklp as String = ""
	Dim ketklp as String = ""

	Dim SQL0 As String
	Dim SQL1 As String
	Dim SQL2 As String
	Dim SQLjo As String
	
	Dim statusexist as Integer = 0
	
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
	Dim maxi as Integer = 0
	Dim maxii as Integer = 0
	
	Dim namaklp as String = ""

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


	If NOT Request.QueryString("namaklp") is Nothing Then
		 namaklp = Request.QueryString("namaklp")
	End If
	
response.write(namaklp)	
	
'***************************************************************************************************************	

	aaa = 1
    Try
        If (aaa=1)
%>
<!--#include file="koneksi.aspx"-->
<%				
			
			myCmd = myConn.CreateCommand

'==========================================================================
' Cari Daftar Divisi
'==========================================================================
			a = 0
			SQL0 = "SELECT * FROM [klp] WHERE klp_nama='" & namaklp & "'"
response.write(SQL0)
			myCmd.CommandText = SQL0
			myConn.Open()
			myReader = myCmd.ExecuteReader()
            If myReader.HasRows Then
				a = 1
			End If
			myConn.Close()
response.write(a)			
        End If

    Catch ex As Exception
        'lbl1.Text="Error while connecting Server."
		'Response.Write (ex.Message)

    Finally
		'myReader.Close()
        'myConn.Close() 'Whether there is error or not. Close the connection.

    End Try
%>
