<%@ Page Language="VB" %>
<%@ Import Namespace="System" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data.OleDb" %> 
<%@ Import Namespace="System.Configuration" %>
<%@ Import Namespace="System.Web.Configuration" %>
 
<script runat="server">

</script>

<%
    'Inherits System.Windows.Forms.Form
    'Create ADO.NET objects.
    Dim myConn As SqlConnection
    Dim myCmd As SqlCommand
	Dim myCmdi As SqlCommand
    Dim myReader As SqlDataReader
    'results As String
	Dim a, aaa, jmla, jmldata as integer

	Dim b as Single = 0
	Dim c as Integer = 0
	Dim idx as Integer = 0
	Dim userid as Integer = 0
	Dim kcx as String = ""
	Dim statval as Integer = 0

	Dim SQL0 As String
	
	Dim statuskc as Integer = 0
	Dim statuskcp as Integer = 0

	If NOT Request.QueryString("kode") is Nothing Then
		kcx = Request.QueryString("kode")
	End If

%>
<!--#include file="koneksi.aspx"-->
<%
	aaa = 1
    Try
        If aaa=1
			myCmd = myConn.CreateCommand
			SQL0 = "SELECT * FROM [divisi] WHERE (div_nama='" & kcx & "')"
			myCmd.CommandText = SQL0
			myConn.Open()
			myReader = myCmd.ExecuteReader()
            If myReader.HasRows Then
				statval = 1
			End If
			myConn.Close()	
			response.write(statval)
        End If

    Catch ex As Exception
        'lbl1.Text="Error while connecting Server."
		'Response.Write (ex.Message)

    Finally
		'myReader.Close()
        'myConn.Close() 'Whether there is error or not. Close the connection.

    End Try
%>

