
<!-- saved from url=(0044)http://www.bayuajie.com/demo-menu/accord.php -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<link rel="stylesheet" type="text/css" href="./sidebarmenu/style-accord.css" media="screen">
 <script src="./sidebarmenu/jquery-latest.min.js" type="text/javascript"></script>
 <script src="./sidebarmenu/script.js"></script>
 



 
</head>
<body>


<div id="cssmenu1">
<ul>
	<li><a href="main.aspx" style="padding-top:8px; text-shadow: rgba(0, 0, 0, 0.35) 0px 1px 1px;"><span style="border-color: rgba(0, 0, 0, 0.35);">HOME</span>&nbsp;&nbsp;&nbsp;&nbsp;<img src="./images/home1.png" alt="" style="vertical-align: bottom; padding-top:auto;width:16px;height:16px;"></a>

	</li>
<%
	'If Val(suratq)=1 OR Val(levelaksesq) > 1 Then
	If Val(suratq)=1 OR Val(levelaksesq) > 1 Then
%>
	<li class="has-sub"><a style="text-shadow: rgba(0, 0, 0, 0.35) 0px 1px 1px;"><span style="border-color: rgba(0, 0, 0, 0.35);">PENOMORAN SURAT</span><span class="holder" style="border-color: rgba(0, 0, 0, 0.35);"></span></a>
		<ul style="display: none;">
			<li><a id="various5" href="surat.aspx"><span>Surat</span></a></li>
			<li><a id="various6" href="memo.aspx"><span>Memo</span></a></li>
			<li><a id="various7" href="notaintern.aspx"><span>Nota Intern</span></a></li>
<%
'If namaklpq.ToUpper() = "BUM" Then
If Val(sekreq)=1 Then
%>
			<li><a id="various8" href="suratdinas.aspx?jenis=suratkeluar"><span>Surat Keluar</span></a></li>
			<li><a id="various9" href="suratdinas.aspx?jenis=suratjalan"><span>Surat Jalan</span></a></li>
			<li><a id="various10" href="suratdinas.aspx?jenis=surattugas"><span>Surat Tugas</span></a></li>
<%
End If
%>
		</ul>
   </li>
<%
	End If
	
	If Val(sekreq)=1 OR Val(levelaksesq) > 1 Then
	'If Val(sekreq)=1 Then
%>
   <li class="has-sub"><a style="text-shadow: rgba(0, 0, 0, 0.35) 0px 1px 1px;"><span style="border-color: rgba(0, 0, 0, 0.35);">KESEKRETARIATAN</span><span class="holder" style="border-color: rgba(0, 0, 0, 0.35);"></span></a>
		<ul style="display: none;">
			<li><a href="suratsekre.aspx?jenis=SURAT MASUK" id="various21"><span>Surat Masuk</span></a></li>
            <li><a href="suratsekre.aspx?jenis=SURAT KELUAR" id="various22"><span>Surat Keluar</span></a></li>
			<li><a href="suratsekre.aspx?jenis=MEMO MASUK" id="various23"><span>Memo Masuk</span></a></li>
            <li><a href="suratsekre.aspx?jenis=MEMO KELUAR" id="various24"><span>Memo Keluar</span></a></li>
			<li><a href="suratsekre.aspx?jenis=NOTA INTERN MASUK" id="various25"><span>Nota Intern Masuk</span></a></li>
            <li><a href="suratsekre.aspx?jenis=NOTA INTERN KELUAR" id="various26"><span>Nota Intern Keluar</span></a></li>
		</ul>
   </li>
<%
	End If
	
	If Val(sekreq) = 1 Then
%>
   <li class="has-sub"><a style="text-shadow: rgba(0, 0, 0, 0.35) 0px 1px 1px;"><span style="border-color: rgba(0, 0, 0, 0.35);">REG PEMBUKUAN</span><span class="holder" style="border-color: rgba(0, 0, 0, 0.35);"></span></a>
		<ul style="display: none;">
			<li><a href="regvoucher.aspx" id="various21"><span>Reg Voucher</span></a></li>
            <li><a href="regpp.aspx" id="various22"><span>Reg PP</span></a></li>
		</ul>
   </li>
<%
	End If
	If Val(sekreq) = 1 AND Val(levelaksesq) > 1 Then
%>
   <li class="has-sub"><a style="text-shadow: rgba(0, 0, 0, 0.35) 0px 1px 1px;"><span style="border-color: rgba(0, 0, 0, 0.35);">ASET</span><span class="holder" style="border-color: rgba(0, 0, 0, 0.35);"></span></a>
		<ul style="display: none;">
			<li><a href="aset.aspx" id="various21"><span>Daftar Aset</span></a></li>
            <li><a href="asetbaru.aspx" id="various22"><span>Penambahan Aset</span></a></li>
		</ul>
   </li>
<%
   End If
%>
<%   
	'End If
	
	If Val(adaq)=1 Then
		'If namaklpq.ToUpper() = "OPB" Then
%>

   <li class="has-sub"><a href="pengadaan.aspx" style="text-shadow: rgba(0, 0, 0, 0.35) 0px 1px 1px;"><span style="border-color: rgba(0, 0, 0, 0.35);">PENGADAAN</span><span class="holder" style="border-color: rgba(0, 0, 0, 0.35);"></span></a>
		<ul style="">
			<li><a href="permohonan.aspx" id="various31"><span>Permohonan</span></a></li>
			<!--li><a href="evaluasikelengkapan.aspx" id="various32"><span>Evaluasi Kelengkapan</span></a></li>
			<li><a href="prosesacc.aspx" id="various33"><span>Proses ACC</span></a></li>
			<li><a href="panggilvendor.aspx" id="various34"><span>Panggil Vendor</span></a></li>
			<li><a href="anwijzing.aspx" id="various35"><span>Aanwijzing</span></a></li>
			<li><a href="negoharga.aspx" id="various36"><span>Nego Harga</span></a></li>
			<li><a href="ujikepatuhan.aspx" id="various37"><span>Checklist Uji Kepatuhan</span></a></li>
			<li><a href="suratpenunjukan.aspx" id="various38"><span>Surat Penunjukan</span></a></li>
			<li><a href="spk.aspx" id="various39"><span>SPK</span></a></li>
			<li><a href="pks.aspx" id="various40"><span>PKS</span></a></li-->
			<li><a href="proses.aspx" id="various41"><span>Proses</span></a></li>
			<li><a id="various6" href="skp.aspx"><span>Keputusan (SK)</span></a></li>
			<li><a id="various7" href="pks.aspx"><span>PKS</span></a></li>
			<!--li><a href="selesai.aspx" id="various41"><span>Selesai</span></a></li-->
		</ul>
	</li>
<%
		'End If
	End If
%>
    <!--li class="has-sub"><a style="text-shadow: rgba(0, 0, 0, 0.35) 0px 1px 1px;"><span style="border-color: rgba(0, 0, 0, 0.35);">ASET</span><span class="holder" style="border-color: rgba(0, 0, 0, 0.35);"></span></a>
            <ul style="display: none;">
				<li><a href="aset.aspx" id="various51"><span>Daftar Aset</span></a></li>				
			</ul>
	</li-->
    <li class="has-sub"><a style="text-shadow: rgba(0, 0, 0, 0.35) 0px 1px 1px;"><span style="border-color: rgba(0, 0, 0, 0.35);">DATA</span><span class="holder" style="border-color: rgba(0, 0, 0, 0.35);"></span></a>
		<ul style="display: none;">
			<li><a href="user.aspx" id="various61"><span>User</span></a></li>
<%
'If Val(levelaksesq) > 1 AND namaklpq.ToUpper() = "BOC" Then
If Val(levelaksesq) > 1 AND Val(suratq)=1 Then
%>
			<li><a href="divisi.aspx" id="various62"><span>Divisi</span></a></li>
			<li><a href="kelompok.aspx" id="various63"><span>Kelompok</span></a></li>			
			<li><a href="rubrik.aspx" id="various64"><span>Rubrik</span></a></li>
<%
End If
%>
		</ul>
	</li>
	<li><a href="logout.aspx" style="padding-top:6px;text-shadow: rgba(0, 0, 0, 0.35) 0px 3px 1px;"><span style="border-color: rgba(0, 0, 0, 0.35);">Logout</span>&nbsp;&nbsp;&nbsp;&nbsp;<img src="./images/logout22.png" alt="" style="vertical-align: bottom; padding-top:2px;width:20px;height:20px;"></a>

	</li>
</ul>
</div>

</body></html>