<!DOCTYPE html>

<%@ Page Language="VB" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data.OleDb" %> 
<%@ Import Namespace="System.Configuration" %>
<%@ Import Namespace="System.Web.Configuration" %>
 
<script runat="server">

</script>

<html>
<head>
    <meta charset="utf-8">
    <!-- Somehow I got an error, so I comment the title, just uncomment to show -->
    <!-- <title>Animated Sidebar Menu | CodingLab</title> -->
    <link rel="stylesheet" href="style.css">
    <script src="https://kit.fontawesome.com/a076d05399.js"></script>
<!-- Font Awesome JS -->
<script src="https://kit.fontawesome.com/b99e675b6e.js"></script>
<!-- jQuery -->
<script src="https://code.jquery.com/jquery-3.4.1.min.js"></script>

    <!--script src="./js/a076d05399.js"></script>
	<script src="./js/b99e675b6e.js"></script>
	<script src="./js/jquery-3.4.1.min.js"></script-->

<!--link rel="stylesheet" href="css1/styles.css"-->

<link href="./datepicker4/dcalendar.picker.css" rel="stylesheet" type="text/css">

<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">


<style>
.example-element {
	margin-top: 0;
	position: fixed;
	background-color: #0A9DE7;
	width: 230px;
	height: 100%;
}
</style>

<style>
body {
  margin: 0;
  font-family: Arial, Helvetica, sans-serif;
}

.topnav {
  overflow: hidden;
  background-color: #0A9DE7;
  //position: relative;
  margin-top: 0;
}

.topnav a {
  float: left;
  display: block;
  color: #f2f2f2;
  text-align: center;
  padding: 17px 8px;
  margin-top: 5px;
  text-decoration: none;
  font-size: 17px;
}

.topnav a:hover {
  //background-color: #ddd;
  //color: black;
}

.topnav a.active {
  background-color: #4CAF50;
  color: white;
}

.topnav .icon {
  display: none;
}

@media screen and (max-width: 600px) {
  .topnav a:not(:first-child) {display: none;}
  .topnav a.icon {
    float: left;
    display: block;
  }
}

@media screen and (max-width: 600px) {
  .topnav.responsive {position: relative;}
  .topnav.responsive .icon {
    position: absolute;
    right: 0;
    top: 0;
  }
  .topnav.responsive a {
    float: none;
    display: block;
    text-align: left;
  }
}
</style>

<style>
.container {
  position: relative;
  text-align: left;
  color: white;
  font-family: Calibri;
  font-size: 0.9em;
}

.welcome {
    left: 20px;
    line-height: 100px;
    margin-top: -40px;
	margin-left: 0px;
    position: absolute;
    text-align: left;
    //top: 50%;
    //width: 100%;
}
.userin {
    left: 0px;
    line-height: 100px;
	margin-top: -5px;
    margin-left: 70px;
    position: absolute;
    text-align: left;
	color: yellow;
    //top: 50%;
    //width: 100%;
}
.posisi {
    left: 0px;
    line-height: 100px;
	margin-top: 20px;
    margin-left: 70px;
    position: absolute;
    text-align: left;
	font-size: 0.8em;
	color: #D7F5FF;
}
.photo {
    left: 0px;
    line-height: 100px;
	margin-top: 20px;
    margin-left: 20px;
    position: absolute;
    text-align: left;
	color: yellow;
    //top: 50%;
    //width: 100%;
}
</style>

<style>
.button, .button:visited {
background-color: #222;
display: inline-block;
padding: 5px 10px 6px;
color: #fff;
text-decoration: none;
-moz-border-radius: 6px;
-webkit-border-radius: 6px;
-moz-box-shadow: 0 1px 3px rgba(0,0,0,0.6);
-webkit-box-shadow: 0 1px 3px rgba(0,0,0,0.6);
text-shadow: 0 -1px 1px rgba(0,0,0,0.25);
border-bottom: 1px solid rgba(0,0,0,0.25);
position: relative;
cursor: pointer;
font-family: Calibri;

}
.button:hover {
background-color: #111;
color: #fff;
} .button:active {
top: 1px;
} 
.small.button, .small.button:visited {
font-size: 0.8em;
padding: 8px 14px 6px;
} 
.button, .button:visited,
.medium.button, .medium.button:visited {
font-size: 1.2em;
font-family: calibri;
font-weight: normal;
line-height: 1;
text-shadow: 0 -1px 1px rgba(0,0,0,0.25); 
} 
.medium.button, .medium.button:visited {
font-size: 1.1em;
padding: 6px 14px 8px;
} 
.large.button, .large.button:visited {
font-size: 1.1em;
padding: 18px 22px 16px;
} 
.pink.button, .magenta.button:visited {
background-color: #e22092;
} 
.pink.button:hover {
background-color: #c81e82;
} 
.green.button, .green.button:visited {
background-color: #00AC04;
} 
.green.button:hover {
background-color: #749a02;
} 
.red.button, .red.button:visited {
background-color: #e62727;
}
.red.button:hover {
background-color: #cf2525;
}
.orange.button, .orange.button:visited {
background-color: #ff5c00;
} 
.orange.button:hover {
background-color: #d45500;
}
.blue.button, .blue.button:visited {
background-color: #2981e4;
}
.blue.button:hover {
background-color: #2575cf;
}
.yellow.button, .yellow.button:visited {
background-color: #ffb515;
}
.yellow.button:hover {
background-color: #fc9200;
} 
</style>
<style> 
.textbox { 
    border: 1px solid #c4c4c4; 
    height: 26px; 
    width: 240px; 
    font-size: 13px; 
    padding: 4px 4px 4px 4px; 
    border-radius: 4px; 
    -moz-border-radius: 4px; 
    -webkit-border-radius: 4px; 
    box-shadow: 0px 0px 4px #d9d9d9; 
    -moz-box-shadow: 0px 0px 4px #d9d9d9; 
    -webkit-box-shadow: 0px 0px 4px #d9d9d9; 
} 
 
.textbox:focus { 
    outline: none; 
    border: 1px solid #7bc1f7; 
    box-shadow: 0px 0px 4px #7bc1f7; 
    -moz-box-shadow: 0px 0px 4px #7bc1f7; 
    -webkit-box-shadow: 0px 0px 4px #7bc1f7; 
</style>
<style>
.textboxnarrow { 
    border: 1px solid #c4c4c4; 
    height: 26px; 
    width: 50px; 
    font-size: 13px; 
    padding: 4px 4px 4px 4px; 
    border-radius: 4px; 
    -moz-border-radius: 4px; 
    -webkit-border-radius: 4px; 
    box-shadow: 0px 0px 4px #d9d9d9; 
    -moz-box-shadow: 0px 0px 4px #d9d9d9; 
    -webkit-box-shadow: 0px 0px 4px #d9d9d9; 
} 
 
.textboxnarrow:focus { 
    outline: none; 
    border: 1px solid #7bc1f7; 
    box-shadow: 0px 0px 4px #7bc1f7; 
    -moz-box-shadow: 0px 0px 4px #7bc1f7; 
    -webkit-box-shadow: 0px 0px 4px #7bc1f7; 
</style>
<style>
.textboxmid { 
    border: 1px solid #c4c4c4; 
    height: 26px; 
    width: 130px; 
    font-size: 13px; 
    padding: 4px 4px 4px 4px; 
    border-radius: 4px; 
    -moz-border-radius: 4px; 
    -webkit-border-radius: 4px; 
    box-shadow: 0px 0px 4px #d9d9d9; 
    -moz-box-shadow: 0px 0px 4px #d9d9d9; 
    -webkit-box-shadow: 0px 0px 4px #d9d9d9; 
} 
 
.textboxmid:focus { 
    outline: none; 
    border: 1px solid #c4c4c4; 
    box-shadow: 0px 0px 4px #7bc1f7; 
    -moz-box-shadow: 0px 0px 4px #7bc1f7; 
    -webkit-box-shadow: 0px 0px 4px #7bc1f7; 
} 
 </style>
 <style>
.textboxwide { 
    border: 1px solid #c4c4c4; 
    height: 26px; 
    width: 250px; 
    font-size: 13px; 
    padding: 4px 4px 4px 4px; 
    border-radius: 4px; 
    -moz-border-radius: 4px; 
    -webkit-border-radius: 4px; 
    box-shadow: 0px 0px 4px #d9d9d9; 
    -moz-box-shadow: 0px 0px 4px #d9d9d9; 
    -webkit-box-shadow: 0px 0px 4px #d9d9d9; 
} 
 
.textboxwide:focus { 
    outline: none; 
    border: 1px solid #7bc1f7; 
    box-shadow: 0px 0px 4px #7bc1f7; 
    -moz-box-shadow: 0px 0px 4px #7bc1f7; 
    -webkit-box-shadow: 0px 0px 4px #7bc1f7; 
} 
 </style>
 <style> 
  .textboxtgl { 
    border: 1px solid #c4c4c4; 
    height: 26px; 
    width: 160px; 
    font-size: 13px; 
    padding: 4px 4px 4px 4px; 
    border-radius: 2px; 
    -moz-border-radius: 2px; 
    -webkit-border-radius: 2px; 
    box-shadow: 0px 0px 2px #d9d9d9; 
    -moz-box-shadow: 0px 0px 2px #d9d9d9; 
    -webkit-box-shadow: 0px 0px 2px #d9d9d9; 
} 
 
.textboxtgl:focus { 
    outline: none; 
    border: 1px solid #7bc1f7; 
    box-shadow: 0px 0px 2px #7bc1f7; 
    -moz-box-shadow: 0px 0px 2px #7bc1f7; 
    -webkit-box-shadow: 0px 0px 2px #7bc1f7; 
} 
 </style>
 
<style>
.inp-icon{
background: url(cal4.png)no-repeat 100%;
background-size: 16px;
 }
</style>

<style>
.pagination {
  display: inline-block;
}

.pagination a {
  color: black;
  float: left;
  padding: 4px 12px;
  text-decoration: none;
  border: 1px solid #ddd;
  border-radius: 3px;
  font-size: 14px;
}

.pagination a.active {
  background-color: #C5C5E2;
  color: white;
  border: 1px solid #C5C5E2;
}

.pagination a:hover:not(.active) {background-color: #ddd;}

.pagination a:first-child {
  border-top-left-radius: 5px;
  border-bottom-left-radius: 5px;
}

.pagination a:last-child {
  border-top-right-radius: 5px;
  border-bottom-right-radius: 5px;
}
</style>

<style>
/* -------------------- Page Styles (not required) */
div { margin: 0px; }

/* -------------------- Select Box Styles: bavotasan.com Method (with special adaptations by ericrasch.com) */
/* -------------------- Source: http://bavotasan.com/2011/style-select-box-using-only-css/ */
.styled-select {
   background: url(15xvbd5.png) no-repeat 96% 0;
   height: 26px;
   overflow: hidden;
   width: 240px;
}

.styled-select select {
   background: transparent;
   border: none;
   font-size: 14px;
   height: 26px;
   padding: 5px; /* If you add too much padding here, the options won't show in IE */
   width: 268px;
}

.styled-select.slate {
   background: url(http://i62.tinypic.com/2e3ybe1.jpg) no-repeat right center;
   height: 30px;
   width: 240px;
}

.styled-select.slate select {
   border: 1px solid #ccc;
   font-size: 16px;
   height: 30px;
   width: 268px;
}

/* -------------------- Rounded Corners */
.rounded {
   -webkit-border-radius: 20px;
   -moz-border-radius: 20px;
   border-radius: 20px;
}

.semi-square {
   -webkit-border-radius: 5px;
   -moz-border-radius: 5px;
   border-radius: 5px;
}

/* -------------------- Colors: Background */
.slate   { background-color: #ddd; }
.green   { background-color: #779126; }
.blue    { background-color: #3b8ec2; }
.yellow  { background-color: #eec111; }
.black   { background-color: #000; }

/* -------------------- Colors: Text */
.slate select   { color: #000; }
.green select   { color: #fff; }
.blue select    { color: #fff; }
.yellow select  { color: #000; }
.black select   { color: #fff; }


/* -------------------- Select Box Styles: danielneumann.com Method */
/* -------------------- Source: http://danielneumann.com/blog/how-to-style-dropdown-with-css-only/ */
#mainselection select {
   border: 0;
   color: #EEE;
   background: transparent;
   font-size: 20px;
   font-weight: bold;
   padding: 2px 10px;
   width: 378px;
   *width: 350px;
   *background: #58B14C;
   -webkit-appearance: none;
}

#mainselection {
   overflow:hidden;
   width:350px;
   -moz-border-radius: 9px 9px 9px 9px;
   -webkit-border-radius: 9px 9px 9px 9px;
   border-radius: 9px 9px 9px 9px;
   box-shadow: 1px 1px 11px #330033;
   background: #58B14C url("15xvbd5.png") no-repeat scroll 319px center;
}


/* -------------------- Select Box Styles: stackoverflow.com Method */
/* -------------------- Source: http://stackoverflow.com/a/5809186 */
select#soflow, select#soflow-color {
   -webkit-appearance: button;
   -webkit-border-radius: 2px;
   -webkit-box-shadow: 0px 1px 3px rgba(0, 0, 0, 0.1);
   -webkit-padding-end: 20px;
   -webkit-padding-start: 2px;
   -webkit-user-select: none;
   background-image: url(15xvbd5.png), -webkit-linear-gradient(#FAFAFA, #F4F4F4 40%, #E5E5E5);
   background-position: 99% center;
   background-repeat: no-repeat;
   border: 1px solid #AAA;
   color: #555;
   font-size: 13px;
   margin: 2px;
   overflow: hidden;
   padding: 2px 8px;
   text-overflow: ellipsis;
   white-space: nowrap;
   width: 90px;
}

select#soflow-color {
   color: #fff;
   background-image: url(15xvbd5.png), -webkit-linear-gradient(#779126, #779126 40%, #779126);
   background-color: #779126;
   -webkit-border-radius: 20px;
   -moz-border-radius: 20px;
   border-radius: 20px;
   padding-left: 10px;
}
select#soflow1, select#soflow-color {
   -webkit-appearance: button;
   -webkit-border-radius: 2px;
   -webkit-box-shadow: 0px 1px 3px rgba(0, 0, 0, 0.1);
   -webkit-padding-end: 20px;
   -webkit-padding-start: 2px;
   -webkit-user-select: none;
   background-image: url(15xvbd5.png), -webkit-linear-gradient(#FAFAFA, #F4F4F4 40%, #E5E5E5);
   background-position: 99% center;
   background-repeat: no-repeat;
   border: 1px solid #AAA;
   color: #555;
   font-size: 13px;
   margin: 2px;
   overflow: hidden;
   padding: 2px 8px;
   text-overflow: ellipsis;
   white-space: nowrap;
   width: 150px;
}

select#soflow1-color {
   color: #fff;
   background-image: url(15xvbd5.png), -webkit-linear-gradient(#779126, #779126 40%, #779126);
   background-color: #779126;
   -webkit-border-radius: 20px;
   -moz-border-radius: 20px;
   border-radius: 20px;
   padding-left: 10px;
}
select#soflow1A, select#soflow-color {
   -webkit-appearance: button;
   -webkit-border-radius: 2px;
   -webkit-box-shadow: 0px 1px 3px rgba(0, 0, 0, 0.1);
   -webkit-padding-end: 20px;
   -webkit-padding-start: 2px;
   -webkit-user-select: none;
   background-image: url(15xvbd5.png), -webkit-linear-gradient(#FAFAFA, #F4F4F4 40%, #E5E5E5);
   background-position: 99% center;
   background-repeat: no-repeat;
   border: 1px solid #AAA;
   color: #555;
   font-size: 13px;
   margin: 2px;
   overflow: hidden;
   padding: 2px 8px;
   text-overflow: ellipsis;
   white-space: nowrap;
   width: 150px;
}

select#soflow1A-color {
   color: #fff;
   background-image: url(15xvbd5.png), -webkit-linear-gradient(#779126, #779126 40%, #779126);
   background-color: #779126;
   -webkit-border-radius: 20px;
   -moz-border-radius: 20px;
   border-radius: 20px;
   padding-left: 10px;
}
select#soflow11, select#soflow11-color {
   -webkit-appearance: button;
   -webkit-border-radius: 2px;
   -webkit-box-shadow: 0px 1px 3px rgba(0, 0, 0, 0.1);
   -webkit-padding-end: 20px;
   -webkit-padding-start: 2px;
   -webkit-user-select: none;
   background-image: url(15xvbd5.png), -webkit-linear-gradient(#FAFAFA, #F4F4F4 40%, #E5E5E5);
   background-position: 99% center;
   background-repeat: no-repeat;
   border: 1px solid #AAA;
   color: #555;
   font-size: 13px;
   margin: 0px;
   overflow: hidden;
   padding: 2px 8px;
   text-overflow: ellipsis;
   white-space: nowrap;
   width: 120px;
}

select#soflow11-color {
   color: #fff;
   background-image: url(15xvbd5.png), -webkit-linear-gradient(#779126, #779126 40%, #779126);
   background-color: #779126;
   -webkit-border-radius: 20px;
   -moz-border-radius: 20px;
   border-radius: 20px;
   padding-left: 10px;
}
select#soflow12, select#soflow12-color {
   -webkit-appearance: button;
   -webkit-border-radius: 2px;
   -webkit-box-shadow: 0px 1px 3px rgba(0, 0, 0, 0.1);
   -webkit-padding-end: 20px;
   -webkit-padding-start: 2px;
   -webkit-user-select: none;
   background-image: url(15xvbd5.png), -webkit-linear-gradient(#FAFAFA, #F4F4F4 40%, #E5E5E5);
   background-position: 99% center;
   background-repeat: no-repeat;
   border: 1px solid #AAA;
   color: #555;
   font-size: 13px;
   margin: 0px;
   overflow: hidden;
   padding: 4px 8px;
   text-overflow: ellipsis;
   white-space: nowrap;
   width: 150px;
}

select#soflow12-color {
   color: #fff;
   background-image: url(15xvbd5.png), -webkit-linear-gradient(#779126, #779126 40%, #779126);
   background-color: #779126;
   -webkit-border-radius: 20px;
   -moz-border-radius: 20px;
   border-radius: 20px;
   padding-left: 10px;
}
select#soflow13, select#soflow13-color {
   -webkit-appearance: button;
   -webkit-border-radius: 2px;
   -webkit-box-shadow: 0px 1px 3px rgba(0, 0, 0, 0.1);
   -webkit-padding-end: 20px;
   -webkit-padding-start: 2px;
   -webkit-user-select: none;
   background-image: url(15xvbd5.png), -webkit-linear-gradient(#FAFAFA, #F4F4F4 40%, #E5E5E5);
   background-position: 99% center;
   background-repeat: no-repeat;
   border: 1px solid #AAA;
   color: #555;
   font-size: 13px;
   margin: 0px;
   overflow: hidden;
   padding: 4px 8px;
   text-overflow: ellipsis;
   white-space: nowrap;
   width: 150px;
}

select#soflow13-color {
   color: #fff;
   background-image: url(15xvbd5.png), -webkit-linear-gradient(#779126, #779126 40%, #779126);
   background-color: #779126;
   -webkit-border-radius: 20px;
   -moz-border-radius: 20px;
   border-radius: 20px;
   padding-left: 10px;
}
select#soflow14, select#soflow14-color {
   -webkit-appearance: button;
   -webkit-border-radius: 2px;
   -webkit-box-shadow: 0px 1px 3px rgba(0, 0, 0, 0.1);
   -webkit-padding-end: 20px;
   -webkit-padding-start: 2px;
   -webkit-user-select: none;
   background-image: url(15xvbd5.png), -webkit-linear-gradient(#FAFAFA, #F4F4F4 40%, #E5E5E5);
   background-position: 99% center;
   background-repeat: no-repeat;
   border: 1px solid #AAA;
   color: #555;
   font-size: 13px;
   margin: 0px;
   overflow: hidden;
   padding: 4px 8px;
   text-overflow: ellipsis;
   white-space: nowrap;
   width: 150px;
}

select#soflow14-color {
   color: #fff;
   background-image: url(15xvbd5.png), -webkit-linear-gradient(#779126, #779126 40%, #779126);
   background-color: #779126;
   -webkit-border-radius: 20px;
   -moz-border-radius: 20px;
   border-radius: 20px;
   padding-left: 10px;
}
select#soflow15, select#soflow15-color {
   -webkit-appearance: button;
   -webkit-border-radius: 2px;
   -webkit-box-shadow: 0px 1px 3px rgba(0, 0, 0, 0.1);
   -webkit-padding-end: 20px;
   -webkit-padding-start: 2px;
   -webkit-user-select: none;
   background-image: url(15xvbd5.png), -webkit-linear-gradient(#FAFAFA, #F4F4F4 40%, #E5E5E5);
   background-position: 99% center;
   background-repeat: no-repeat;
   border: 1px solid #AAA;
   color: #555;
   font-size: 13px;
   margin: 0px;
   overflow: hidden;
   padding: 4px 8px;
   text-overflow: ellipsis;
   white-space: nowrap;
   width: 150px;
}

select#soflow15-color {
   color: #fff;
   background-image: url(15xvbd5.png), -webkit-linear-gradient(#779126, #779126 40%, #779126);
   background-color: #779126;
   -webkit-border-radius: 20px;
   -moz-border-radius: 20px;
   border-radius: 20px;
   padding-left: 10px;
}
select#soflow16, select#soflow16-color {
   -webkit-appearance: button;
   -webkit-border-radius: 2px;
   -webkit-box-shadow: 0px 1px 3px rgba(0, 0, 0, 0.1);
   -webkit-padding-end: 20px;
   -webkit-padding-start: 2px;
   -webkit-user-select: none;
   background-image: url(15xvbd5.png), -webkit-linear-gradient(#FAFAFA, #F4F4F4 40%, #E5E5E5);
   background-position: 99% center;
   background-repeat: no-repeat;
   border: 1px solid #AAA;
   color: #555;
   font-size: 13px;
   margin: 0px;
   overflow: hidden;
   padding: 4px 8px;
   text-overflow: ellipsis;
   white-space: nowrap;
   width: 150px;
}

select#soflow16-color {
   color: #fff;
   background-image: url(15xvbd5.png), -webkit-linear-gradient(#779126, #779126 40%, #779126);
   background-color: #779126;
   -webkit-border-radius: 20px;
   -moz-border-radius: 20px;
   border-radius: 20px;
   padding-left: 10px;
}
select#soflow17, select#soflow17-color {
   -webkit-appearance: button;
   -webkit-border-radius: 2px;
   -webkit-box-shadow: 0px 1px 3px rgba(0, 0, 0, 0.1);
   -webkit-padding-end: 20px;
   -webkit-padding-start: 2px;
   -webkit-user-select: none;
   background-image: url(15xvbd5.png), -webkit-linear-gradient(#FAFAFA, #F4F4F4 40%, #E5E5E5);
   background-position: 99% center;
   background-repeat: no-repeat;
   border: 1px solid #AAA;
   color: #555;
   font-size: 13px;
   margin: 0px;
   overflow: hidden;
   padding: 4px 8px;
   text-overflow: ellipsis;
   white-space: nowrap;
   width: 150px;
}

select#soflow17-color {
   color: #fff;
   background-image: url(15xvbd5.png), -webkit-linear-gradient(#779126, #779126 40%, #779126);
   background-color: #779126;
   -webkit-border-radius: 20px;
   -moz-border-radius: 20px;
   border-radius: 20px;
   padding-left: 10px;
}
select#soflow18, select#soflow18-color {
   -webkit-appearance: button;
   -webkit-border-radius: 2px;
   -webkit-box-shadow: 0px 1px 3px rgba(0, 0, 0, 0.1);
   -webkit-padding-end: 20px;
   -webkit-padding-start: 2px;
   -webkit-user-select: none;
   background-image: url(15xvbd5.png), -webkit-linear-gradient(#FAFAFA, #F4F4F4 40%, #E5E5E5);
   background-position: 99% center;
   background-repeat: no-repeat;
   border: 1px solid #AAA;
   color: #555;
   font-size: 13px;
   margin: 0px;
   overflow: hidden;
   padding: 4px 8px;
   text-overflow: ellipsis;
   white-space: nowrap;
   width: 150px;
}

select#soflow18-color {
   color: #fff;
   background-image: url(15xvbd5.png), -webkit-linear-gradient(#779126, #779126 40%, #779126);
   background-color: #779126;
   -webkit-border-radius: 20px;
   -moz-border-radius: 20px;
   border-radius: 20px;
   padding-left: 10px;
}
select#soflow21, select#soflow21-color {
   -webkit-appearance: button;
   -webkit-border-radius: 2px;
   -webkit-box-shadow: 0px 1px 3px rgba(0, 0, 0, 0.1);
   -webkit-padding-end: 20px;
   -webkit-padding-start: 2px;
   -webkit-user-select: none;
   background-image: url(15xvbd5.png), -webkit-linear-gradient(#FAFAFA, #F4F4F4 40%, #E5E5E5);
   background-position: 99% center;
   background-repeat: no-repeat;
   border: 1px solid #AAA;
   color: #555;
   font-size: 13px;
   margin: 0px;
   overflow: hidden;
   padding: 4px 8px;
   text-overflow: ellipsis;
   white-space: nowrap;
   width: 150px;
}

select#soflow21-color {
   color: #fff;
   background-image: url(15xvbd5.png), -webkit-linear-gradient(#779126, #779126 40%, #779126);
   background-color: #779126;
   -webkit-border-radius: 20px;
   -moz-border-radius: 20px;
   border-radius: 20px;
   padding-left: 10px;
}
select#soflow22, select#soflow22-color {
   -webkit-appearance: button;
   -webkit-border-radius: 2px;
   -webkit-box-shadow: 0px 1px 3px rgba(0, 0, 0, 0.1);
   -webkit-padding-end: 20px;
   -webkit-padding-start: 2px;
   -webkit-user-select: none;
   background-image: url(15xvbd5.png), -webkit-linear-gradient(#FAFAFA, #F4F4F4 40%, #E5E5E5);
   background-position: 99% center;
   background-repeat: no-repeat;
   border: 1px solid #AAA;
   color: #555;
   font-size: 13px;
   margin: 0px;
   overflow: hidden;
   padding: 4px 8px;
   text-overflow: ellipsis;
   white-space: nowrap;
   width: 150px;
}

select#soflow22-color {
   color: #fff;
   background-image: url(15xvbd5.png), -webkit-linear-gradient(#779126, #779126 40%, #779126);
   background-color: #779126;
   -webkit-border-radius: 20px;
   -moz-border-radius: 20px;
   border-radius: 20px;
   padding-left: 10px;
}
</style>

<style>
.css-button-3 {
	font-size: 12px;
	border-radius: 5px;
	border: solid 0px #18a418;
	color: #ffffff;
	background: linear-gradient(180deg, #32cd32 5%, #18a418 100%);
	box-shadow: 0px 5px 4px -2px #616174;
	//font-family: Arial;
	cursor: pointer;
	text-align: center;
	user-select: none;
	display: inline-flex;
	justify-content: center;
	align-items: center;
}
.css-button-3:hover {
	background: linear-gradient(180deg, #18a418 5%, #32cd32 100%);
}
.css-button-3:active {
	position: relative;
	top: 1px;
}
.css-button-3 > span {
	display: block;
}
.css-button-3-icon {
	padding: 6px 8px;
	border-right: 1px solid rgba(255, 255, 255, 0.16);
	box-shadow: rgba(0, 0, 0, 0.14) -1px 0px 0px inset;
}
.css-button-3-text {
	padding: 6px 8px;
}
</style>

<style>
.css-button-4 {
	font-size: 12px;
	border-radius: 5px;
	border: solid 0px #18a418;
	color: #ffffff;
	background: linear-gradient(180deg, #6085FF 5%, #0548a4 100%);
	box-shadow: 0px 5px 4px -2px #616174;
	//font-family: Arial;
	cursor: pointer;
	text-align: center;
	user-select: none;
	display: inline-flex;
	justify-content: center;
	align-items: center;
}
.css-button-4:hover {
	background: linear-gradient(180deg, #0548a4 5%, #6085FF 100%);
}
.css-button-4:active {
	position: relative;
	top: 1px;
}
.css-button-4 > span {
	display: block;
}
.css-button-4-icon {
	padding: 6px 8px;
	border-right: 1px solid rgba(255, 255, 255, 0.16);
	box-shadow: rgba(0, 0, 0, 0.14) -1px 0px 0px inset;
}
.css-button-4-text {
	padding: 6px 8px;
}
</style>

<style>
.container1 {
  margin: 20px 0px auto;
}
/* Ignore above this line */ 
 
input[type="checkbox"] {
  display: none;
}

input[type="checkbox"] + label span,input[type="checkbox"]:checked + label span {
  display: inline-block;
  width: 24px;
  height: 24px;
  margin: -20px 8px 0 0;
  vertical-align: top;
  cursor: pointer;
}

input[type="checkbox"] + label span {
  background: url('./images/variable-unchecked73.png') no-repeat;
}

input[type="checkbox"]:checked + label span {
  background: url('./images/variable-checked73.png') no-repeat;
} 
</style>

<style>

.container2 {
    margin: 0px 0px;
    width: 80px;
    text-align: left;
}
.container2 > .switch {
    display: block;
    margin: 0px 0px;
}
.switch {
    position: relative;
    display: inline-block;
    vertical-align: top;
    width: 62px;
    height: 20px;
    padding: 3px;
    background-color: white;
    border-radius: 18px;
    box-shadow: inset 0 -1px white, inset 0 1px 1px rgba(0, 0, 0, 0.05);
    cursor: pointer;
    background-image: -webkit-linear-gradient(top, #eeeeee, white 25px);
    background-image: -moz-linear-gradient(top, #eeeeee, white 25px);
    background-image: -o-linear-gradient(top, #eeeeee, white 25px);
    background-image: linear-gradient(to bottom, #eeeeee, white 25px);
}
.switch-input {
    position: absolute;
    top: 0;
    left: 0;
    opacity: 0;
}
.switch-label {
    position: relative;
    display: block;
    height: inherit;
    font-size: 10px;
    text-transform: uppercase;
    background: #eceeef;
    border-radius: inherit;
    box-shadow: inset 0 1px 2px rgba(0, 0, 0, 0.12), inset 0 0 2px rgba(0, 0, 0, 0.15);
    -webkit-transition: 0.15s ease-out;
    -moz-transition: 0.15s ease-out;
    -o-transition: 0.15s ease-out;
    transition: 0.15s ease-out;
    -webkit-transition-property: opacity background;
    -moz-transition-property: opacity background;
    -o-transition-property: opacity background;
    transition-property: opacity background;
}
.switch-label:before,
.switch-label:after {
    position: absolute;
    top: 50%;
    margin-top: -.5em;
    line-height: 1;
    -webkit-transition: inherit;
    -moz-transition: inherit;
    -o-transition: inherit;
    transition: inherit;
}
.switch-label:before {
    content: attr(data-off);
    right: 11px;
    color: #aaa;
    text-shadow: 0 1px rgba(255, 255, 255, 0.5);
}
.switch-label:after {
    content: attr(data-on);
    left: 11px;
    color: white;
    text-shadow: 0 1px rgba(0, 0, 0, 0.2);
    opacity: 0;
}
.switch-input:checked ~ .switch-label {
    background: #47a8d8;
    box-shadow: inset 0 1px 2px rgba(0, 0, 0, 0.15), inset 0 0 3px rgba(0, 0, 0, 0.2);
}
.switch-input:checked ~ .switch-label:before {
    opacity: 0;
}
.switch-input:checked ~ .switch-label:after {
    opacity: 1;
}
.switch-handle {
    position: absolute;
    top: 4px;
    left: 4px;
    width: 18px;
    height: 18px;
    background: white;
    border-radius: 10px;
    box-shadow: 1px 1px 5px rgba(0, 0, 0, 0.2);
    background-image: -webkit-linear-gradient(top, white 40%, #f0f0f0);
    background-image: -moz-linear-gradient(top, white 40%, #f0f0f0);
    background-image: -o-linear-gradient(top, white 40%, #f0f0f0);
    background-image: linear-gradient(to bottom, white 40%, #f0f0f0);
    -webkit-transition: left 0.15s ease-out;
    -moz-transition: left 0.15s ease-out;
    -o-transition: left 0.15s ease-out;
    transition: left 0.15s ease-out;
}
.switch-handle:before {
    content: '';
    position: absolute;
    top: 50%;
    left: 50%;
    margin: -6px 0 0 -6px;
    width: 12px;
    height: 12px;
    background: #f9f9f9;
    border-radius: 6px;
    box-shadow: inset 0 1px rgba(0, 0, 0, 0.02);
    background-image: -webkit-linear-gradient(top, #eeeeee, white);
    background-image: -moz-linear-gradient(top, #eeeeee, white);
    background-image: -o-linear-gradient(top, #eeeeee, white);
    background-image: linear-gradient(to bottom, #eeeeee, white);
}
.switch-input:checked ~ .switch-handle {
    left: 40px;
    box-shadow: -1px 1px 5px rgba(0, 0, 0, 0.2);
}
.switch-green > .switch-input:checked ~ .switch-label {
    background: #4fb845;
}

</style>

<style>
/* The container3 */
.container3 {
  display: block;
  position: relative;
  padding-left: 35px;
  margin-bottom: 12px;
  cursor: pointer;
  -webkit-user-select: none;
  -moz-user-select: none;
  -ms-user-select: none;
  user-select: none;
}

/* Hide the browser's default checkbox */
.container3 input {
  position: absolute;
  opacity: 0;
  cursor: pointer;
  height: 0;
  width: 0;
}

/* Create a custom checkbox */
.checkmark {
  position: absolute;
  top: 0;
  left: 0;
  height: 20px;
  width: 20px;
  background-color: #eee;
}

/* On mouse-over, add a grey background color */
.container3:hover input ~ .checkmark {
  background-color: #ccc;
}

/* When the checkbox is checked, add a blue background */
.container3 input:checked ~ .checkmark {
  background-color: #2196F3;
}

/* Create the checkmark/indicator (hidden when not checked) */
.checkmark:after {
  content: "";
  position: absolute;
  display: none;
}

/* Show the checkmark when checked */
.container3 input:checked ~ .checkmark:after {
  display: block;
}

/* Style the checkmark/indicator */
.container3 .checkmark:after {
  left: 6px;
  top: 3px;
  width: 5px;
  height: 10px;
  border: solid white;
  border-width: 0 3px 3px 0;
  -webkit-transform: rotate(45deg);
  -ms-transform: rotate(45deg);
  transform: rotate(45deg);
}
</style>


<script type="text/javascript">     
function PreviewImage(no) 
{       
	var namaim = "uploadImage" + no;
	var uploadim = "uploadPreview" + no;  
	var oFReader = new FileReader();
	var namafil = document.getElementById(namaim).value;
	var namafilearr = namafil.split(".");
	var noarr = namafilearr.length;
	var jenisfile = namafilearr[noarr-1];

//alert(namafil);	
	oFReader.readAsDataURL(document.getElementById(namaim).files[0]);
	
	if (jenisfile==='pdf')
	{
		oFReader.onload = function (oFREvent) {             
		document.getElementById(uploadim).src = './images/pdficon1.png'; 	
		};
	}
	else
	{
		oFReader.onload = function (oFREvent) {             
		document.getElementById(uploadim).src = oFREvent.target.result; 		
		};
	}
	//document.forms["memobarufrm"].submit();
	//memousulanbaru();

} 
</script>

<script>
function xxxhapussurat(id)
{
	if (confirm("Apakah anda ingin hapus data ?") == true) 
	{
		var tenure = prompt("Silahkan masukkan kode hapus", "");
		
		if (tenure == "opr1234") {
			var outx;
			var linkx = "hapuskctgo.aspx?id="+id;
			xhttp = new XMLHttpRequest();
			xhttp.onreadystatechange = function() {
				if (this.readyState == 4 && this.status == 200) {
					//document.getElementById("txtHint").innerHTML = this.responseText;
					outx = this.responseText;
					if (outx==1 || outx=="1")
					{
						alert("Data sudah dihapus.");
						window.location.reload();
					}
				}
			};
			xhttp.open("GET", linkx, true);
			xhttp.send();
		}
	}
}
</script>

<script>
function hapususer(id)
{
	if (confirm("Apakah anda ingin hapus data ?") == true) 
	{
		var tenure = prompt("Silahkan masukkan kode hapus", "");
		
		if (tenure == "123") {
			document.getElementsByName("deluserid")[0].value = id;
			document.forms["delfrm"].submit();
		}
	}
}
</script>

<script>
function clearparam()
{
	document.getElementsByName("perihal")[0].value = "";
	document.getElementsByName("tglperiode1")[0].value = "";
	document.getElementsByName("tglperiode2")[0].value = "";
	document.forms["filterfrm"].submit();
}
</script>

<script>
function viewpage2(nohal)
{
var myval = document.getElementsByName("namanpp")[0].value;
if (myval==null || myval=="")
{
	var halink = "surat.aspx?page=" + nohal + "&cari=none";
}
else
{
	var halink = "surat.aspx?page=" + nohal + "&cari=" + myval;
}
//window.open(halink);
window.location.href = halink;
	//document.getElementsByName("pageno")[0].value = nohal;
	//document.forms["pagefrm"].submit();
}
</script>

<script>
function viewpage(nohal, totpage)
{
	if (nohal > totpage || nohal < 1)
	{
		return false;
	}
	document.getElementsByName("pagenos")[0].value = nohal;
	document.getElementsByName("statusfilter")[0].value = 1;
	document.forms["filterfrm"].submit();
}
</script>

<script>
function viewfilter()
{
//var myval = document.getElementsByName("namanpp")[0].value;
//var halink = "main.aspx?page=1&cari=" + myval;
//window.open(halink);
//window.location.href = halink;
	document.getElementsByName("statusfilter")[0].value = 1;
	document.getElementsByName("pagenos")[0].value = 1;
	document.forms["filterfrm"].submit();
}
</script>

<script>
function clearparam()
{
//var myval = document.getElementsByName("namanpp")[0].value;
//var halink = "main.aspx?page=1&cari=" + myval;
//window.open(halink);
//window.location.href = halink;
	document.getElementsByName("namanpp")[0].value = "";
	document.getElementsByName("divisi")[0].value = "";
	document.getElementsByName("statusfilter")[0].value = 0;
	document.forms["filterfrm"].submit();
}
</script>

<script>
function databaru()
{
	document.getElementById("databarutab").style.display = "block";
	document.getElementById("dataedittab").style.display = "none";
	document.getElementById("filtertab").style.display = "none";
	document.getElementById("barubtntab").style.display = "none";
	document.getElementById("datatab").style.display = "none";
	document.getElementById("totalpagetab").style.display = "none";
	document.getElementById("pagingtab").style.display = "none";
	//document.getElementById("clearbtn").style.display = "none";
	//document.getElementById("filtertab").style.display = "none";
}
</script>

<script>
function hidedatabaru()
{
	document.getElementById("databarutab").style.display = "none";
	document.getElementById("dataedittab").style.display = "none";
	document.getElementById("filtertab").style.display = "block";
	document.getElementById("barubtntab").style.display = "block";
	document.getElementById("datatab").style.display = "block";
	document.getElementById("totalpagetab").style.display = "block";
	document.getElementById("pagingtab").style.display = "block";
	//document.getElementById("clearbtn").style.display = "none";
	//document.getElementById("filtertab").style.display = "none";
}
</script>

<script>
function hidedataedit(levelakses, pengadaan)
{
	document.getElementById("databarutab").style.display = "none";
	document.getElementById("dataedittab").style.display = "none";
	//if (levelakses>1 && kel=="BUM")
	if (levelakses>1 && pengadaan==0)
	{
		document.getElementById("filtertab").style.display = "block";
		document.getElementById("barubtntab").style.display = "block";
	}
	//if (levelakses>1 && kel!="BUM")
	if (levelakses>1 && pengadaan==1)
	{
		document.getElementById("filtertab").style.display = "block";
	}

	document.getElementById("datatab").style.display = "block";
	document.getElementById("totalpagetab").style.display = "block";
	document.getElementById("pagingtab").style.display = "block";
	//document.getElementById("clearbtn").style.display = "none";
	//document.getElementById("filtertab").style.display = "none";
}
</script>

<script>
function hideeditbaru()
{
	document.getElementById("databarutab").style.display = "none";
	document.getElementById("dataedittab").style.display = "none";
	document.getElementById("filtertab").style.display = "block";
	document.getElementById("barubtntab").style.display = "block";
	document.getElementById("datatab").style.display = "block";
	document.getElementById("totalpagetab").style.display = "block";
	document.getElementById("pagingtab").style.display = "block";
	//document.getElementById("clearbtn").style.display = "none";
	//document.getElementById("filtertab").style.display = "none";
}
</script>

<script>
function getkel(dival)
{
	var outx;
	var linkx = "getklp.aspx?divid="+dival;
	xhttp = new XMLHttpRequest();
	xhttp.onreadystatechange = function() {
		if (this.readyState == 4 && this.status == 200) {
			//document.getElementById("txtHint").innerHTML = this.responseText;
			outx = this.responseText;
			document.getElementById("soflow13").innerHTML = outx;
			document.getElementById("soflow13").disabled = false;
			//document.getElementById("languageList").value = outx;
			//document.getElementsByName("barunama")[0].value = outarr[1];	
		}
	};
	xhttp.open("GET", linkx, true);
	xhttp.send();
}
</script>

<script>
function getunit(kelval)
{
	var outx;
	var linkx = "getunit.aspx?klpid="+kelval;

	xhttp = new XMLHttpRequest();
	xhttp.onreadystatechange = function() {
		if (this.readyState == 4 && this.status == 200) {
			//document.getElementById("txtHint").innerHTML = this.responseText;
			outx = this.responseText;
			document.getElementById("soflow14").innerHTML = outx;
			document.getElementById("soflow14").disabled = false;
			//document.getElementById("languageList").value = outx;
			//document.getElementsByName("barunama")[0].value = outarr[1];	
		}
	};
	xhttp.open("GET", linkx, true);
	xhttp.send();
}
</script>

<script>
function savebaru()
{
	var nppku = document.getElementsByName("barunpp")[0].value;
	var suratchk = document.getElementById("barusuratchk");
	var sekrechk = document.getElementById("barusekrechk");
	var adachk = document.getElementById("baruadachk");
	var asetchk = document.getElementById("baruasetchk");
	var pwd1 = document.getElementsByName("barupwd")[0].value;
	var pwd2 = document.getElementsByName("barupwd2")[0].value;
	var suratval = 0;
	var sekreval = 0;
	var adaval = 0;
	var asetval = 0;
	if (suratchk.checked==true)
	{
		suratval = 1;
	}
	if (sekrechk.checked==true)
	{
		sekreval = 1;
	}
	if (adachk.checked==true)
	{
		adaval = 1;
	}
	if (asetchk.checked==true)
	{
		asetval = 1;
	}

	if (pwd1 !== pwd2 || pwd1.length < 4 || pwd1.length > 20)
	{
		alert("Password anda tidak valid/konsisten. Harap masukkan password yang sama dan 4-20 karakter.");
		return false;
	}

	document.getElementsByName("barucsurat")[0].value = suratval;
	document.getElementsByName("barucsekre")[0].value = sekreval;
	document.getElementsByName("barucada")[0].value = adaval;
	document.getElementsByName("barucaset")[0].value = asetval;
		
		//document.getElementsByName("baruju")[0].value = 1;
		//document.getElementById("baruasetchk").checked = true;
//---------------------------------------------------------------------		
//Cek apakah user dgn NPP sudah ada
	var outx;
	var linkx = "ceknpp.aspx?npp="+nppku;
	xhttp = new XMLHttpRequest();
	xhttp.onreadystatechange = function() {
		if (this.readyState == 4 && this.status == 200) {
			//document.getElementById("txtHint").innerHTML = this.responseText;
			outx = this.responseText;
			if (outx==1 || outx=="1")
			{
				alert("Data dengan NPP yang sama sudah ada. Silahkan periksa.");
				return false;
			}
			else
			{
				document.forms["barufrm"].submit();
			}
		}
	};
	xhttp.open("GET", linkx, true);
	xhttp.send();		
//---------------------------------------------------------------------		
		
	
}
</script>

<script>
function saveedit()
{
	var nppku = document.getElementsByName("editnpp")[0].value;
	var suratchk = document.getElementById("editsuratchk");
	var sekrechk = document.getElementById("editsekrechk");
	var adachk = document.getElementById("editadachk");
	var asetchk = document.getElementById("editasetchk");
	var pwd1 = document.getElementsByName("editpwd")[0].value;
	var pwd2 = document.getElementsByName("editpwd2")[0].value;
	var suratval = 0;
	var sekreval = 0;
	var adaval = 0;
	var asetval = 0;
	if (suratchk.checked==true)
	{
		suratval = 1;
	}
	if (sekrechk.checked==true)
	{
		sekreval = 1;
	}
	if (adachk.checked==true)
	{
		adaval = 1;
	}
	if (asetchk.checked==true)
	{
		asetval = 1;
	}

	if (pwd1 !== pwd2 || pwd1.length < 4 || pwd1.length > 20)
	{
		alert("Password anda tidak valid/konsisten. Harap masukkan password yang sama dan 4-20 karakter.");
		return false;
	}

	document.getElementsByName("editcsurat")[0].value = suratval;
	document.getElementsByName("editcsekre")[0].value = sekreval;
	document.getElementsByName("editcada")[0].value = adaval;
	document.getElementsByName("editcaset")[0].value = asetval;
	document.forms["editfrm"].submit();
}
</script>

<script>
function editdatax(id, nama, npp, divi, kelo, unit, posisi, level, menusurat, menusekre, menuada, menuaset, userlevel, userposisi)
{
	document.getElementById("databarutab").style.display = "none";
	document.getElementById("dataedittab").style.display = "block";
	if (!document.getElementById("filtertab")==false)
	{
		document.getElementById("filtertab").style.display = "none";
	}
	if (!document.getElementById("barubtntab")==false)
	{
		document.getElementById("barubtntab").style.display = "none";
	}
	document.getElementById("datatab").style.display = "none";
	document.getElementById("totalpagetab").style.display = "none";
	document.getElementById("pagingtab").style.display = "none";

	document.getElementsByName("editnama")[0].value = nama;
	document.getElementsByName("editnpp")[0].value = npp;
	document.getElementById("soflow15").value = divi;
	document.getElementById("soflow16").value = kelo;
	document.getElementById("soflow17").value = unit;
	document.getElementById("soflow18").value = posisi;
	document.getElementById("soflow22").value = level;
	document.getElementsByName("editpwd")[0].value = "11111111";
	document.getElementsByName("editpwd2")[0].value = "11111111";
	document.getElementsByName("editid")[0].value = id;
	
	if (menusurat==1)
	{
		document.getElementById("editsuratchk").checked=true;
	}
	if (menusekre==1)
	{
		document.getElementById("editsekrechk").checked=true;
	}
	if (menuada==1)
	{
		document.getElementById("editadachk").checked=true;
	}
	if (menuaset==1)
	{
		document.getElementById("editasetchk").checked=true;
	}
	
	if (userlevel==1)
	{
		document.getElementsByName("editnama")[0].disabled = true;
		document.getElementsByName("editnpp")[0].disabled = true;
		document.getElementById("soflow15").disabled = true;
		document.getElementById("soflow16").disabled = true;
		document.getElementById("soflow17").disabled = true;
		document.getElementById("soflow18").disabled = true;
		document.getElementById("soflow22").disabled = true;
		document.getElementById("editsuratchk").disabled = true;
		document.getElementById("editsekrechk").disabled = true;
		document.getElementById("editadachk").disabled = true;
		document.getElementById("editasetchk").disabled = true;
	}

}
</script>

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
	Dim a, aaa, b, bb, jmla, jmldata as Long
	Dim user0 as string
	Dim pwd0 as string
	Dim namauserx as String
	Dim aksesx as String
	Dim rubrikx(1000) as String
	
	Dim idq As String = ""
	Dim namauserq As String = ""
	Dim statusq As String = ""
	Dim nppq as String = ""
	Dim jabatanq As String = ""
	Dim klpq As String = ""
	Dim unitq As String = ""
	Dim posisiq As String = ""
	Dim suratq As String = ""
	Dim sekreq As String = ""
	Dim adaq as String
	Dim asetq as String
	Dim statusaktifq as String = ""
	Dim divq As String = ""
	Dim levelaksesq as String = ""
	Dim levelaksesx as Integer = 0
	Dim posisix as String = ""
	
	Dim idxx as String
	Dim nosuratxx as String
	Dim tglxx as String
	Dim jenisxx as String
	Dim rubrikxx as String
	Dim perihalxx as String
	Dim tujuanxx as String
	Dim tembusanxx as String
	Dim jenisundanganxx as String
	Dim cuserxx as String
	Dim namafilexx as String
	Dim namafolderxx as String
	
	'Dim pageno as Integer = 1
	Dim rowperpage as Integer = 20
	Dim pageno as Integer = 1
	Dim startpage as Long = 0
	Dim endpage as Long = 0
	Dim jmlpage as Integer = 1
	Dim pagerows as Integer = 20
	Dim rowsperpage as Integer = 20
	Dim hal as Integer = 1
	'Dim b as Single = 0
	Dim c as Long = 0
	Dim useridx as Integer = 0
	Dim itop as Integer = 1
	Dim ii as Long = 0
	Dim k as Long = 0
	Dim userid as Integer = 1
	Dim jmldiv as Integer = 0
	Dim jmlkel as Integer = 0
	Dim jmlunit as Integer = 0
	Dim jmlposisi as Integer = 0
	Dim jmllevelakses as Integer = 0
	Dim namadivx(1000) as String
	Dim namakelx(1000) as String
	Dim namaunitx(1000) as String
	Dim kelidx(1000) as String
	Dim unitidx(1000) as String

	Dim SQL0 As String
	Dim SQL1 As String
	Dim SQL2 As String
	
	Dim startval as Integer = 0
	Dim endval as Integer = 0
	Dim datastrx(5000) as String
	Dim datastrxx as String
	Dim namanppr as String = ""
	Dim divisir as String = ""
	Dim namanppstr as String = ""
	Dim divstr as String = ""
	Dim tgl1 as String = ""
	Dim tgl2 as String = ""
	Dim tgl10 as string = ""
	Dim tgl20 as String = ""
	Dim tglstr as String = ""
	Dim tglsu as String
	Dim statusf as Integer = 0
	Dim namaklpq as String = ""
	Dim i as Long
	Dim l as Integer
	Dim m as integer
	Dim n as Integer
	Dim jmlrubrik as Integer = 0
	
	Dim setmenu as String = "surat"
	
	Dim tahuniki as Integer
	Dim tahunwingi as Integer
	Dim tahunsesuk as Integer
	
	Dim idz(2000) as String
	Dim namaz(2000) as String
	Dim nppz(2000) as String
	Dim menusuratz(2000) as String
	Dim menusekz(2000) as String
	Dim menuadaz(2000) as String
	Dim menuasetz(2000) as String
	Dim divz(2000) as String
	Dim klpz(2000) as String
	Dim unitz(2000) as String
	Dim divzi(2000) as String
	Dim klpzi(2000) as String
	Dim unitzi(2000) as String
	Dim dividx(100) as String
	Dim posisiz(2000) as String
	Dim posisizi(2000) as String
	Dim levelz(2000) as String
	Dim posisiidx(200) as String
	Dim posisistrx(200) as String
	Dim levelidx(20) as String
	Dim levelstrx(20) as String

	'lbl1.Text = Hour(Now) & ":" & Minute(Now) & ":" & Second(Now)
	tahuniki = Val(Year(Now))
	tahunwingi = tahuniki - 1
	tahunsesuk = tahuniki + 1
	
	aaa=1
	
	hal = 1
	pageno = 1
	'GET
	'If NOT Request.QueryString("pageno") is Nothing Then
	'	 pageno = Request.QueryString("pageno")
	'	 hal = Cint(pageno)
	'End If
	

    Try
        If aaa=1 Then
			'myConn = New SqlConnection("Data Source=DESKTOP-S7N8DF9;Initial Catalog=KCT;User ID=sa;Password=sa;")
			pagerows = (hal-1) * 20
%>

<!--#include file="koneksi.aspx"-->

<%			
			If Not Request.Form("pagenos") is Nothing Then
				pageno = Val(Request.Form("pagenos"))
				'hal = pageno
				'hal = Cint(pageno)
			End If
			If pageno=0 Then
				pageno = 1
			End If
'response.write("---------------------------------------------" & hal & "<br>")
			If NOT Request.Form("namanpp") is Nothing Then
				namanppr = Request.Form("namanpp")
			End If
			If NOT Request.Form("divisi") is Nothing Then
				divisir = Request.Form("divisi")
			End If
			If NOT Request.Form("statusfilter") is Nothing Then
				statusf = Cint(Request.Form("statusfilter"))
			End If
			
			user0 = Session("username")
			idq = Session("userid")
			nppq = Session("idemployee")
			If not Session("namauser") is nothing Then
				namauserq = Session("namauser")
			End If
			If not Session("namaklp") is nothing Then
				namaklpq = Session("namaklp")
			End If
			jabatanq = Session("jabatan")
			klpq = Session("klp")
			unitq = Session("unit")
			posisiq = Session("posisi")
			posisix = Session("posisistr")
			suratq = Session("menusurat")
			sekreq = Session("menusekre")
			adaq = Session("menuada")
			asetq = Session("menuaset")
			divq = Session("divisi")
			levelaksesq = Session("levelakses")
			levelaksesx = Val(levelaksesq)
		
			If namauserq.length()<2 Then
%>
<script>
alert('Anda tidak beraktivitas terlalu lama (idle). Silahkan signin lagi');
window.top.location.href = "logout.aspx"; 
</script>
<%
			End If
			
			myCmd = myConn.CreateCommand
	
'If perihal IsNot Nothing || NOT perihal Is Nothing
			namanppstr = ""
			If namanppr.length()>0 Then
				namanppstr = " AND (nama LIKE '%" & namanppr & "%' OR idemployee LIKE '%" & namanppr & "%')"
			End If
			
			If divisir.length() > 0 Then
				divstr = " AND [user].div=" & divisir & ""
			End If

'============================================================================
' Data user jika levelakses userlogin adalah USER
'============================================================================
			If levelaksesx <= 1 Then
				namanppstr  = " AND (idemployee='" & user0 & "')"
			End If


'============================================================================
' Data USER
'============================================================================
			'SQL0 = "SELECT ID FROM [user] WHERE (ID>0" & perihalstr & tglstr & ") ORDER BY tglsurat DESC"
			SQL1 = "SELECT * FROM [user] ORDER BY nama ASC"
					
			SQL1 = "SELECT * FROM [user] WHERE (ID>0" & namanppstr & divstr & ") ORDER BY nama ASC"
'response.write("-------------------------------------------" & SQL1 & "<br>")
			myCmd.CommandText = SQL1
			myConn.Open()
			myReader = myCmd.ExecuteReader()		
			
			a = 0
			jmldata = 0
            If myReader.HasRows Then
                While myReader.Read()
					a = a + 1
					idz(a) = myReader("ID").ToString()
					namaz(a) = myReader("nama").ToString()
					nppz(a) = myReader("idemployee").ToString()
					menusuratz(a) = myReader("menusurat").ToString()
					menusekz(a) = myReader("menusekretaris").ToString()
					menuadaz(a) = myReader("menupengadaan").ToString()
					menuasetz(a) = myReader("menuaset").ToString()
					divzi(a) = myReader("div").ToString()
					klpzi(a) = myReader("klp").ToString()
					unitzi(a) = myReader("unit").ToString()
					posisizi(a) = myReader("posisi").ToString()
					levelz(a) = myReader("levelakses").ToString()
					
					If menusuratz(a).length() < 1 Then
						menusuratz(a) = 0
					End If
					If menusekz(a).length() < 1 Then
						menusekz(a) = 0
					End If
					If menuadaz(a).length() < 1 Then
						menuadaz(a) = 0
					End If
					If menuasetz(a).length() < 1 Then
						menuasetz(a) = 0
					End If
					If posisizi(a).length() < 1 Then
						posisizi(a) = 0
					End If
				End While
			End If
			myConn.Close()
			jmldata = a

			For a=1 to jmldata
'--------------------------------------------------------------------------
' Cari Nama Unit
'--------------------------------------------------------------------------
				SQL0 = "SELECT * FROM [unit] WHERE u_id=" & unitzi(a)
				unitz(a) = 0
				myCmd.CommandText = SQL0
				myConn.Open()
				myReader = myCmd.ExecuteReader()
				If myReader.HasRows Then
					While myReader.Read()
						unitz(a) = myReader("u_nama").ToString()
					End While
				End If
				myConn.Close()
'--------------------------------------------------------------------------
' Cari Nama Kelompok
'--------------------------------------------------------------------------
				SQL0 = "SELECT * FROM [klp] WHERE klp_id=" & klpzi(a)
				myCmd.CommandText = SQL0
				myConn.Open()
				myReader = myCmd.ExecuteReader()
				If myReader.HasRows Then
					While myReader.Read()
						klpz(a) = myReader("klp_nama").ToString()
					End While
				End If
				myConn.Close()	
'--------------------------------------------------------------------------
' Cari Nama Divisi
'--------------------------------------------------------------------------
				SQL0 = "SELECT * FROM [divisi] WHERE div_id=" & divzi(a)
				myCmd.CommandText = SQL0
				myConn.Open()
				myReader = myCmd.ExecuteReader()
				If myReader.HasRows Then
					While myReader.Read()
						divz(a) = myReader("div_nama").ToString()
					End While
				End If
				myConn.Close()
'--------------------------------------------------------------------------
' Cari Nama Posisi
'--------------------------------------------------------------------------
				SQL0 = "SELECT * FROM [posisi] WHERE po_id=" & posisizi(a)
				myCmd.CommandText = SQL0
				myConn.Open()
				myReader = myCmd.ExecuteReader()
				If myReader.HasRows Then
					While myReader.Read()
						posisiz(a) = myReader("posisi").ToString()
					End While
				End If
				myConn.Close()				
				
			Next a

'==========================================================================
' Cari Daftar Divisi
'==========================================================================
			SQL0 = "SELECT * FROM [divisi] ORDER BY div_nama ASC"
			myCmd.CommandText = SQL0
			myConn.Open()
			myReader = myCmd.ExecuteReader()
			b = 0
			jmldiv = 0
            If myReader.HasRows Then
                While myReader.Read()
					b = b + 1
					dividx(b) = myReader("div_id").ToString()
					namadivx(b) = myReader("div_nama").ToString()
				End While
			End If
			myConn.Close()
			jmldiv = b
'response.write("------------------------------------------x-" & jmldata)
'==========================================================================
' Cari Daftar Kelompok
'==========================================================================
			SQL0 = "SELECT * FROM [klp] ORDER BY klp_nama ASC"
			myCmd.CommandText = SQL0
			myConn.Open()
			myReader = myCmd.ExecuteReader()
			b = 0
            If myReader.HasRows Then
                While myReader.Read()
					b = b + 1
					kelidx(b) = myReader("klp_id").ToString()
					namakelx(b) = myReader("klp_nama").ToString()
				End While
			End If
			myConn.Close()
			jmlkel = b

'==========================================================================
' Cari Daftar Unit
'==========================================================================
			SQL0 = "SELECT * FROM [unit] ORDER BY u_nama ASC"
			myCmd.CommandText = SQL0
			myConn.Open()
			myReader = myCmd.ExecuteReader()
			b = 0
            If myReader.HasRows Then
                While myReader.Read()
					b = b + 1
					unitidx(b) = myReader("u_id").ToString()
					namaunitx(b) = myReader("u_nama").ToString()
				End While
			End If
			myConn.Close()
			jmlunit = b
'==========================================================================
' Cari Daftar Posisi
'==========================================================================
			SQL0 = "SELECT * FROM [posisi] ORDER BY posisi ASC"
			myCmd.CommandText = SQL0
			myConn.Open()
			myReader = myCmd.ExecuteReader()
			b = 0
            If myReader.HasRows Then
                While myReader.Read()
					b = b + 1
					posisiidx(b) = myReader("po_id").ToString()
					posisistrx(b) = myReader("posisi").ToString()
				End While
			End If
			myConn.Close()
			jmlposisi = b
'==========================================================================
' Cari Daftar Level Akses
'==========================================================================
			SQL0 = "SELECT * FROM [levelakses] ORDER BY jb_id ASC"
			myCmd.CommandText = SQL0
			myConn.Open()
			myReader = myCmd.ExecuteReader()
			b = 0
            If myReader.HasRows Then
                While myReader.Read()
					b = b + 1
					levelidx(b) = myReader("jb_id").ToString()
					levelstrx(b) = myReader("jb_nama").ToString()
				End While
			End If
			myConn.Close()
			jmllevelakses = b
'==========================================================================

'** Cari jumlah halaman
			if (jmldata >= rowperpage) Then
				b = jmldata / rowsperpage
				c = jmldata / rowsperpage
				jmlpage = c
				if (b > c) then
					jmlpage = c + 1
				End If
			else
				jmlpage = 1
			End If

			If hal = jmlpage Then
				rowsperpage = jmldata - ((jmlpage - 1)*20)
			End If

			hal = pageno
			startval = ((hal-1)*20) + 1
			endval = startval + rowsperpage - 1
'*********
'response.write("------------------------------------------x-" & jmldata & "--" & jmlpage & "++++" & hal & "<br>")
'response.write("------------------------------------------x-" & startval & "<br>")

        End If

    Catch ex As Exception
        'lbl1.Text="Error while connecting Server."
		'Response.Write (ex.Message)

    Finally
		'myReader.Close()
        'myConn.Close() 'Whether there is error or not. Close the connection.

    End Try
%>

<div class="example-element"></div>
 
<!--#include file="menu.aspx"-->

<!--#include file="topbar.aspx"-->

<div class="container">
			<div class="welcome">Selamat Datang,</div>
			<div class="photo"><img src="./images/user5.png" alt="" style="width:40px;height:40px;"></div>
			<div class="userin"><i><b><% Response.Write(namauserq) %></b></i></div>
			<div class="posisi" STYLE="font-size: 0.8em"><i><% Response.Write(posisix) %></i></div>
</div>

<table style="text-align: left; margin-left: 270px; margin-right: auto; margin-top: auto;" width="860px" border="0" cellspacing="0" cellpadding="0">
	<tr height="30px">
		<td colspan="4" style="padding-left: 20px; text-align: left; font-size: 1.0em">
		</td>
	</tr>
	<tr height="20px">
		<td colspan="4" style="padding-left: 0px; text-align: left; font-size: 1.5em">
			<FONT face="Arial" color="#666666">DAFTAR USER</FONT>
		</td>
	</tr>
	<tr height="30px">
		<td colspan="4" style="padding-left: 20px; text-align: left; font-size: 1.0em">
		
		</td>
	</tr>
</table>

<form method="post" name="barufrm" id="barufrm" action="useraddgo.aspx" enctype="multipart/form-data">
<table id="databarutab" style="text-align: left; margin-left: 270px; margin-right: auto; margin-top: 0; display: none" width="860px" border="0" cellspacing="0" cellpadding="0">
	<tr height="0px">
		<td colspan="4" style="padding-left: 20px; text-align: left; font-size: 1.0em">
		</td>
	</tr>
	<tr height="20px">
		<td colspan="4" style="padding-left: 0px; text-align: left; font-size: 1.3em">
			<FONT face="Arial" color="#44AA44">>> USER BARU</FONT>
		</td>
	</tr>
	<tr height="40px">
		<td colspan="4" style="padding-left: 20px; text-align: left; font-size: 1.0em">
		</td>
	</tr>
	<tr height="35px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: middle" width="180px">
			<FONT face="Arial" color="#000">
				Nama
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 0.9em; vertical-align: middle" width="30px">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: middle" width="550px">
			<FONT face="Arial" color="#000">
				<input style="font-family: Arial; font-size: 0.9em; width: 600px; height: 18px;" type="text" id="barunama" name="barunama" class="textbox" onchange="getnama(this.value);" tabindex="1">	
			</FONT>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.0em" width="100px">										
		</td>
	</tr>
	<tr height="35px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: middle" width="180px">
			<FONT face="Arial" color="#000">
				NPP
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 0.9em; vertical-align: middle" width="30px">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: middle" width="550px">
			<FONT face="Arial" color="#000">
				<input style="font-family: Arial; font-size: 0.9em; width: 130px; height: 18px;" type="text" id="barunpp" name="barunpp" class="textbox" onchange="getnama(this.value);" tabindex="2">	
			</FONT>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em" width="100px">										
		</td>
	</tr>

	<tr height="35px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: middle">
			<FONT face="Arial" color="#000">
				Divisi
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 0.9em; vertical-align: middle">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: middle">
			<FONT face="Arial" color="#000">
				<!--input style="font-family: Arial; font-size: 1.0em; width: 560px" type="text" id="barupengirim" name="barupengirim" class="textbox" onchange="getnama(this.value);" tabindex="4"-->
			<select STYLE="color: #000000; background-color: #ffffff; font-size: 0.9em; width: 130px" id="soflow12" name="barudivisi" onchange="getkel(this.value);" data-placeholder="Jenis" tabindex="3"> 
				<option value="" >&nbsp;&nbsp;</option>
<%
			for b=1 to jmldiv
%>
				<option value="<% response.write(dividx(b)) %>" ><% response.write(namadivx(b)) %></option>
<%
			next b
%>
			</select>		
			</FONT>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em">										
		</td>
	</tr>
	<tr height="35px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Kelompok
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: top">
			<FONT face="Arial" color="#000">
				<!--input style="font-family: Arial; font-size: 1.0em; width: 560px" type="text" id="barupengirim" name="barupengirim" class="textbox" onchange="getnama(this.value);" tabindex="4"-->
			<select STYLE="color: #000000; background-color: #ffffff; font-size: 0.9em; width: 130px" id="soflow13" name="barukelompok" onchange="getunit(this.value);" data-placeholder="Jenis" tabindex="4" disabled="disabled"> 
				<option value="" >&nbsp;&nbsp;</option>
<%
			for b=1 to jmlkel
%>
				<option value="<% response.write(kelidx(b)) %>" ><% response.write(namakelx(b)) %></option>
<%
			next b
%>
			</select>		
			</FONT>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em">										
		</td>
	</tr>
	<tr height="35px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Unit
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: top">
			<FONT face="Arial" color="#000">
				<!--input style="font-family: Arial; font-size: 1.0em; width: 560px" type="text" id="barupengirim" name="barupengirim" class="textbox" onchange="getnama(this.value);" tabindex="4"-->
			<select STYLE="color: #000000; background-color: #ffffff; font-size: 0.9em; width: 550px" id="soflow14" name="baruunit" data-placeholder="Jenis" tabindex="5" disabled="disabled"> 
				<option value="" >&nbsp;&nbsp;</option>
<%
			for b=1 to jmlunit
%>
				<option value="<% response.write(unitidx(b)) %>" ><% response.write(namaunitx(b)) %></option>
<%
			next b
%>
			</select>		
			</FONT>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em">										
		</td>
	</tr>
	<tr height="35px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Posisi
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: top">
			<FONT face="Arial" color="#000">
				<!--input style="font-family: Arial; font-size: 1.0em; width: 560px" type="text" id="barupengirim" name="barupengirim" class="textbox" onchange="getnama(this.value);" tabindex="4"-->
			<select STYLE="color: #000000; background-color: #ffffff; font-size: 0.9em; width: 350px" id="soflow12" name="baruposisi" data-placeholder="Jenis" tabindex="6"> 
				<option value="" >&nbsp;&nbsp;</option>
<%
			for b=1 to jmlposisi
%>
				<option value="<% response.write(posisiidx(b)) %>" ><% response.write(posisistrx(b)) %></option>
<%
			next b
%>
			</select>		
			</FONT>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em">										
		</td>
	</tr>
	<tr height="35px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Level Akses
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: top">
			<FONT face="Arial" color="#000">
				<!--input style="font-family: Arial; font-size: 1.0em; width: 560px" type="text" id="barupengirim" name="barupengirim" class="textbox" onchange="getnama(this.value);" tabindex="4"-->
			<select STYLE="color: #000000; background-color: #ffffff; font-size: 0.9em; width: 150px" id="soflow21" name="barulevel" data-placeholder="Level Akses" tabindex="7"> 
				<option value="" >&nbsp;&nbsp;</option>
<%
			for b=1 to jmllevelakses
%>
				<option value="<% response.write(levelidx(b)) %>" ><% response.write(levelstrx(b)) %></option>
<%
			next b
%>
			</select>		
			</FONT>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em">										
		</td>
	</tr>
	<tr height="35px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Akses Menu
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: top; font-size: 0.9em">
			<FONT face="Arial" color="#000">
				<label class="container3">Surat
				  <!--input id="barusuratchk" name="barusuratchk" type="checkbox" checked="checked" -->
				  <input id="barusuratchk" name="barusuratchk" type="checkbox" tabindex="8">
				  <span class="checkmark"></span>
				</label>
				<label class="container3">Kesekretariatan
				  <input id="barusekrechk" name="barusekrechk" type="checkbox" tabindex="9">
				  <span class="checkmark"></span>
				</label>
				<label class="container3">Pengadaan
				  <input id="baruadachk" name="baruadachk" type="checkbox" tabindex="10">
				  <span class="checkmark"></span>
				</label>
				<label class="container3">Aset
				  <input id="baruasetchk" name="baruasetchk" type="checkbox" tabindex="11">
				  <span class="checkmark"></span>
				</label>
			</FONT>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em">										
		</td>
	</tr>
	<tr height="20px">
		<td colspan="4" style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top">										
		</td>
	</tr>
	<tr height="35px">
		<td colspan="4" style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Jika anda ingin mengubah Password, isi kotak dibawah. Minimal 4 karakter, maksimal 20 karakter.
			</FONT>											
		</td>
	</tr>
	<tr height="35px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: middle">
			<FONT face="Arial" color="#000">
				Password
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 0.9em; vertical-align: middle">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: middle">
			<FONT face="Arial" color="#000">
				<input style="background-color:#FFF9B9; font-size: 0.9em; width: 200px; height: 18px;" type="password" id="barupwd" name="barupwd" class="textbox" tabindex="12">
			</FONT>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.0em">										
		</td>
	</tr>
	<tr height="35px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: middle">
			<FONT face="Arial" color="#000">
				Retype Password
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 0.9em; vertical-align: middle">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: middle">
			<FONT face="Arial" color="#000">
				<input style="background-color:#FFF9B9; font-size: 0.9em; width: 200px; height: 18px;" type="password" id="barupwd2" name="barupwd2" class="textbox" tabindex="13">
			</FONT>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em">										
		</td>
	</tr>
	<tr height="40px">
		<td colspan="3" style="padding-left: 20px; text-align: left;">
		</td>
	</tr>
	<tr height="40px">
		<td style="padding-left: 0px; text-align: left;">
			<a id="batalbtn" href="#" style="font-size: 0.9em" class="medium button orange" onclick="hidedatabaru();"><i class="fa fa-close" aria-hidden="true"></i>&nbsp;&nbsp;Tutup</a>
		</td>
		<td colspan="3" style="padding-left: 0px; text-align: left;">
			<a id="simpanbtn" href="#" style="font-size: 0.9em" class="medium button green" onclick="savebaru();"><i class="fa fa-save" aria-hidden="true"></i>&nbsp;&nbsp;Simpan</a>
		</td>
	</tr>
	<tr height="30px">
		<td colspan="4">
			<input type="hidden" id="barucsurat" name="barucsurat" value="0">
			<input type="hidden" id="barucsekre" name="barucsekre" value="0">
			<input type="hidden" id="barucada" name="barucada" value="0">
			<input type="hidden" id="barucaset" name="barucaset" value="0">
			<input type="hidden" id="editfile" name="editfile">
		</td>
	</tr>
</table>
</form>

<%
'response.write("-----------------------------------x--" & levelaksesq & "=" & posisiq)
%>

<form method="post" name="editfrm" id="editfrm" action="usereditgo.aspx" enctype="multipart/form-data">
<table id="dataedittab" style="text-align: left; margin-left: 270px; margin-right: auto; margin-top: 0; display: none" width="1050px" border="0" cellspacing="0" cellpadding="0">
	<tr height="0px">
		<td colspan="4" style="padding-left: 20px; text-align: left; font-size: 1.0em">
		</td>
	</tr>
	<tr height="20px">
		<td colspan="4" style="padding-left: 0px; text-align: left; font-size: 1.3em">
			<FONT face="Arial" color="#045CFF">VIEW / EDIT</FONT>
		</td>
	</tr>
	<tr height="40px">
		<td colspan="4" style="padding-left: 20px; text-align: left; font-size: 1.0em">
		</td>
	</tr>
	<tr height="35px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: middle" width="220px">
			<FONT face="Arial" color="#000">
				Nama
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 0.9em; vertical-align: middle" width="30px">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: middle" width="700px">
			<FONT face="Arial" color="#000">
				<input style="font-family: Arial; font-size: 0.9em; width: 600px; height: 18px;" type="text" id="editnama" name="editnama" class="textbox" tabindex="1">	
			</FONT>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.0em" width="100px">										
		</td>
	</tr>
	<tr height="35px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: middle">
			<FONT face="Arial" color="#000">
				NPP
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 0.9em; vertical-align: middle">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: middle">
			<FONT face="Arial" color="#000">
<%
			If levelaksesx < 2 Then
%>
				<input style="font-family: Arial; font-size: 0.9em; width: 130px; height: 18px;" type="text" id="editnpp" name="editnpp" class="textbox" tabindex="2" disabled>	
<%
			ElseIf levelaksesx > 1 Then
%>
				<input style="font-family: Arial; font-size: 0.9em; width: 130px; height: 18px;" type="text" id="editnpp" name="editnpp" class="textbox" tabindex="2">
<%
			End If
%>
			</FONT>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em">										
		</td>
	</tr>

	<tr height="35px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Divisi
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: top">
			<FONT face="Arial" color="#000">
				<!--input style="font-family: Arial; font-size: 1.0em; width: 560px" type="text" id="editpengirim" name="editpengirim" class="textbox" onchange="getnama(this.value);" tabindex="4"-->
			<select STYLE="color: #000000; background-color: #ffffff; font-size: 0.9em; width: 130px" id="soflow15" name="editdivisi" onchange="getkel(this.value);" data-placeholder="Jenis" tabindex="3"> 
				<option value="" >&nbsp;&nbsp;</option>
<%
			for b=1 to jmldiv
%>
				<option value="<% response.write(dividx(b)) %>" ><% response.write(namadivx(b)) %></option>
<%
			next b
%>
			</select>		
			</FONT>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em">										
		</td>
	</tr>
	<tr height="35px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Kelompok
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: top">
			<FONT face="Arial" color="#000">
				<!--input style="font-family: Arial; font-size: 1.0em; width: 560px" type="text" id="editpengirim" name="editpengirim" class="textbox" onchange="getnama(this.value);" tabindex="4"-->
			<select STYLE="color: #000000; background-color: #ffffff; font-size: 0.9em; width: 130px" id="soflow16" name="editkelompok" onchange="getunit(this.value);" data-placeholder="Jenis" tabindex="4"> 
				<option value="" >&nbsp;&nbsp;</option>
<%
			for b=1 to jmlkel
%>
				<option value="<% response.write(kelidx(b)) %>" ><% response.write(namakelx(b)) %></option>
<%
			next b
%>
			</select>		
			</FONT>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em">										
		</td>
	</tr>
	<tr height="35px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Unit
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: top">
			<FONT face="Arial" color="#000">
				<!--input style="font-family: Arial; font-size: 1.0em; width: 560px" type="text" id="editpengirim" name="editpengirim" class="textbox" onchange="getnama(this.value);" tabindex="4"-->
			<select STYLE="color: #000000; background-color: #ffffff; font-size: 0.9em; width: 550px" id="soflow17" name="editunit" data-placeholder="Jenis" tabindex="5"> 
				<option value="" >&nbsp;&nbsp;</option>
<%
			for b=1 to jmlunit
%>
				<option value="<% response.write(unitidx(b)) %>" ><% response.write(namaunitx(b)) %></option>
<%
			next b
%>
			</select>		
			</FONT>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em">										
		</td>
	</tr>
	<tr height="35px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Posisi
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: top">
			<FONT face="Arial" color="#000">
				<!--input style="font-family: Arial; font-size: 1.0em; width: 560px" type="text" id="editpengirim" name="editpengirim" class="textbox" onchange="getnama(this.value);" tabindex="4"-->
			<select STYLE="color: #000000; background-color: #ffffff; font-size: 0.9em; width: 350px" id="soflow18" name="editposisi" data-placeholder="Jenis" tabindex="6"> 
				<option value="" >&nbsp;&nbsp;</option>
<%
			for b=1 to jmlposisi
%>
				<option value="<% response.write(posisiidx(b)) %>" ><% response.write(posisistrx(b)) %></option>
<%
			next b
%>
			</select>		
			</FONT>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em">										
		</td>
	</tr>
	<tr height="35px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Level Akses
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: top">
			<FONT face="Arial" color="#000">
				<!--input style="font-family: Arial; font-size: 1.0em; width: 560px" type="text" id="barupengirim" name="barupengirim" class="textbox" onchange="getnama(this.value);" tabindex="4"-->
			<select STYLE="color: #000000; background-color: #ffffff; font-size: 0.9em; width: 150px" id="soflow22" name="editlevel" data-placeholder="Level Akses" tabindex="7"> 
				<option value="" >&nbsp;&nbsp;</option>
<%
			for b=1 to jmllevelakses
%>
				<option value="<% response.write(levelidx(b)) %>" ><% response.write(levelstrx(b)) %></option>
<%
			next b
%>
			</select>		
			</FONT>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em">										
		</td>
	</tr>
	<tr height="35px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Akses Menu
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: top; font-size: 0.9em">
			<FONT face="Arial" color="#000">
				<label STYLE="" class="container3">Surat
				  <!--input id="editsuratchk" name="editsuratchk" type="checkbox" checked="checked" -->
				  <input id="editsuratchk" name="editsuratchk" type="checkbox" tabindex="7">
				  <span class="checkmark"></span>
				</label>
				<label class="container3">Kesekretariatan
				  <input id="editsekrechk" name="editsekrechk" type="checkbox" tabindex="8">
				  <span class="checkmark"></span>
				</label>
				<label class="container3">Pengadaan
				  <input id="editadachk" name="editadachk" type="checkbox" tabindex="9">
				  <span class="checkmark"></span>
				</label>
				<label class="container3">Aset
				  <input id="editasetchk" name="editasetchk" type="checkbox" tabindex="10">
				  <span class="checkmark"></span>
				</label>
			</FONT>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em">										
		</td>
	</tr>
	<tr height="20px">
		<td colspan="4" style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top">										
		</td>
	</tr>
	<tr height="35px">
		<td colspan="4" style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Jika anda ingin mengubah Password, isi kotak dibawah. Minimal 4 karakter, maksimal 20 karakter.
			</FONT>											
		</td>
	</tr>
	<tr height="35px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Password
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: top">
			<FONT face="Arial" color="#000">
				<input style="background-color:#FFF9B9; font-size: 0.9em; width: 200px" type="password" id="editpwd" name="editpwd" class="textbox" value="11111111" tabindex="11">
			</FONT>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em">										
		</td>
	</tr>
	<tr height="35px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Retype Password
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: top">
			<FONT face="Arial" color="#000">
				<input style="background-color:#FFF9B9; font-size: 0.9em; width: 200px" type="password" id="editpwd2" name="editpwd2" class="textbox" value="11111111" tabindex="12">
			</FONT>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.0em">										
		</td>
	</tr>
	<tr height="40px">
		<td colspan="3" style="padding-left: 20px; text-align: left;">
		</td>
	</tr>
	<tr height="40px">
		<td style="padding-left: 0px; text-align: left;">
			<a id="batalbtn" href="#" style="font-size: 0.9em" class="medium button orange" onclick="hidedataedit(<% response.write(levelaksesx) %>, '<% response.write(adaq) %>');"><i class="fa fa-close" aria-hidden="true"></i>&nbsp;&nbsp;Tutup</a>
		</td>
		<td colspan="3" style="padding-left: 0px; text-align: left;">
			<a id="simpanbtn" href="#" style="font-size: 0.9em" class="medium button blue" onclick="saveedit();"><i class="fa fa-save" aria-hidden="true"></i>&nbsp;&nbsp;Simpan</a>
		</td>
	</tr>
	<tr height="30px">
		<td colspan="4">
			<input type="hidden" id="editid" name="editid">
			<input type="hidden" id="editcsurat" name="editcsurat">
			<input type="hidden" id="editcsekre" name="editcsekre">
			<input type="hidden" id="editcada" name="editcada">
			<input type="hidden" id="editcaset" name="editcaset">
			<input type="hidden" id="editfile" name="editfile">
		</td>
	</tr>
</table>
</form>

<table style="text-align: left; margin-left: 270px; margin-right: auto; margin-top: auto;" width="950px" border="0" cellspacing="0" cellpadding="0">
	<tr height="0px">
		<td colspan="4" style="padding-left: 0px; text-align: left; font-size: 1.6em">
		</td>
	</tr>
</table>

<%
If Val(levelaksesx) > 1 Then
%>
<form method="post" name="filterfrm" id="filterfrm" action="user.aspx" enctype="multipart/form-data">
<table id="filtertab" style="text-align: left; margin-left: 270px; margin-right: auto; margin-top: auto;" width="1150px" border="0" cellspacing="0" cellpadding="0" bgcolor="#CAFFEB">
	<tr height="20px">
		<td colspan="4" style="padding-left: 20px; text-align: left; font-size: 1.0em">
		</td>
	</tr>
	<tr height="35px">
		<td style="padding-left: 30px; text-align: left; font-size: 0.9em; vertical-align: middle" width="200px">
			<FONT face="Arial" color="#000">
				Divisi
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 0.9em; vertical-align: middle" width="40px">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: middle" width="600px">
			<FONT face="Arial" color="#000">
				<!--input style="font-family: Arial; font-size: 1.0em; width: 560px" type="text" id="barupengirim" name="barupengirim" class="textbox" onchange="getnama(this.value);" tabindex="4"-->
			<select STYLE="color: #000000; background-color: #ffffff; font-size: 0.9em; width: 130px" id="soflow12" name="divisi" data-placeholder="Jenis" tabindex="4"> 
				<option value="" >&nbsp;&nbsp;</option>
<%
			for b=1 to jmldiv
%>
				<option value="<% response.write(dividx(b)) %>" 
<% 	if dividx(b)=divisir Then 
		response.write("selected") 
	End If 
%>><% response.write(namadivx(b)) %></option>
<%
			next b
%>
			</select>		
			</FONT>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em" width="310px">										
		</td>
	</tr>
	<tr height="35px">
		<td style="padding-left: 0px; padding-left: 30px; text-align: left; font-size: 0.9em; vertical-align: middle">
			<FONT face="Arial" color="#000">
				Nama / NPP
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 0.9em; vertical-align: middle">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: middle">
			<FONT face="Arial" color="#000">
				<!--select id="barutpmobil" name="barutpmobil" STYLE="color: #000000; background-color: #ffffff; font-size: 1.2em; width:320px"-->
				<div>
					<!--input style="padding-right: 10px; text-align: left; font-family: Arial; font-size: 1.2em" type="text" id="brgmerk" name="brgmerk" class="textbox"-->
					<input style="padding-right: 10px; text-align: left; font-size: 0.9em; width: 600px; height: 20px" type="text" id="namanpp" name="namanpp" class="textbox" value="<% Response.Write(namanppr) %>">
				</div>
			</FONT>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em">										
		</td>
	</tr>
	<tr height="20px">
		<td colspan="3" style="padding-left: 20px; text-align: left;">
		</td>
	</tr>
	<tr height="40px">
		<td style="padding-left: 30px; text-align: left;">
			<a id="caribtn" href="#" class="medium button blue" onclick="viewfilter();"><i class="fas fa-search" aria-hidden="true"></i>&nbsp;&nbsp;Cari</a>
		</td>
		<td colspan="2" style="padding-left: 30px; text-align: left;">
			<a id="clearbtn" href="#" class="medium button red" onclick="clearparam();">Reset</a>
		</td>
	</tr>
	<tr height="20px">
		<td colspan="3" style="padding-left: 20px; text-align: left;">
			<input type="hidden" id="userid" name="userid" value="">
			<input type="hidden" id="pagenos" name="pagenos">
			<input type="hidden" id="statusfilter" name="statusfilter" value="0">
		</td>
	</tr>
</table>
</form>

<%
End If
%>

<table style="text-align: left; margin-left: 270px; margin-right: auto; margin-top: auto;" width="1150px" border="0" cellspacing="0" cellpadding="0">
	<tr height="30px">
		<td colspan="4" style="padding-left: 20px; text-align: left; font-size: 1.0em">
		</td>
	</tr>
</table>

<%
'If levelaksesx > 1 AND namaklpq.ToUpper() = "BUM" Then
If levelaksesx > 1 Then
%>
<table id="barubtntab" style="margin-left: 270px;" width="1150px" align="center" cellpadding="2" cellspacing="2" border="0" bgcolor="#EEE" >
	<tr height="38px" align="left">
		<td style="font-family: Arial; font-size: 0.9em;text-align: left;" width="120px">
			<!--a id="barubbm" href="" class="small button blue" onclick="addkct(); return false;"><i class="fas fa-plus fa-sm"></i>&nbsp;&nbsp;New</a-->
			<a href="#" style="text-decoration: none; font-size: 1.1em" class="css-button-3" onclick="databaru();">
				<span class="css-button-3-icon"><i class="fa fa-plus" aria-hidden="true" id="i_pos"></i></span>
				<span class="css-button-3-text"><FONT face="Calibri">Data Baru</FONT></span>
			</a>
		</td>
		<td style="font-family: Arial; font-size: 0.9em;text-align: left;" width="680px">
		</td>
		<td style="font-family: Arial; font-size: 0.9em;text-align: left;" width="160px">
			<a href="#" style="text-decoration: none; font-size: 1.1em" class="css-button-4" onclick="exportcsv();">
				<span class="css-button-4-icon"><i class="fa fa-plus" aria-hidden="true" id="i_pos"></i></span>
				<span class="css-button-4-text"><FONT face="Calibri">Export to CSV</FONT></span>
			</a>
		</td>
		<td style="font-family: Arial; font-size: 0.9em;text-align: left;" width="50px">
		</td>
		<td style="font-family: Arial; font-size: 0.9em;text-align: left; vertical-align: bottom" width="140px">
			<FONT face="Arial" color="#000">Jml data:&nbsp;<% response.write(jmldata.ToString("#,#")) %></FONT>
		</td> 		
	</tr>
</table>

<%
End If
'response.write("-------------------------------------------" & startval & "++" & endval & "<br>")
'response.write("-------------------------------------------" & jmldata & "<br>")
			
'response.write("---------------------------------------------------------" & SQL1 & vbcrlf)
'response.write("---------------------------------------------------------" & pageno & "--" & hal & "   start:" & startval & "-" & endval)

%>

<table id="datatab" style="margin-left: 270px;" width="1150px" align="center" cellpadding="2" cellspacing="2" border="0" bgcolor="#FAFAFA" >
	<tr height="42px" align="left" style="background-color:#D5D5D5;color:Black;" >
		<td style="font-family: Arial; font-size: 0.9em;text-align: center;" width="50px"><b>NO</b></td>                        
		<td style="font-family: Arial; font-size: 0.9em;text-align: center;" width="300px"><b>NAMA</b></td>            
		<td style="font-family: Arial; font-size: 0.9em;text-align: center;" width="80px"><b>NPP</b></td>
		<td style="font-family: Arial; font-size: 0.9em;text-align: center;" width="80px"><b>DIVISI</b></td>
		<td style="font-family: Arial; font-size: 0.9em;text-align: center;" width="100px"><b>KELOMPOK</b></td>
		<td style="font-family: Arial; font-size: 0.9em;text-align: center;" width="190px"><b>UNIT</b></td>
		<td style="font-family: Arial; font-size: 0.9em;text-align: center;" width="200px"><b>POSISI</b></td>
		<td style="font-family: Arial; font-size: 0.9em;text-align: center;" width="75px"><b>EDIT</b></td>
<%
		'If levelaksesx > 1 AND namaklpq.ToUpper() = "BUM" Then
		If levelaksesx > 1 Then
%>
		<td style="font-family: Arial; font-size: 0.9em;text-align: center;" width="75px"><b>HAPUS</b></td>
<%
		End If
%>		
	</tr>

<%

	'm = 46 MOD 2
	'response.write("-------------------------------------------------" & m & " start:" & startval & "--" & endval)
Try
	ii = pagerows
	k = 1
	a = endval - startval
	bb = -1
	n = 0
	for i = startval to endval
		ii = ii + k
		bb = bb + 1
		
		'SQL2 = "SELECT * FROM jenissurat WHERE (js_id=" & jenisxx & ")"	
		'myCmd.CommandText = SQL2
		'myConn.Open()
		'myReader = myCmd.ExecuteReader()
		
		'Dim jenissuratxx AS String=""
		'a = i
		'If myReader.HasRows Then
		'	While myReader.Read()
		'		jenissuratxx = myReader("jenissurat").ToString()
		'	End While
		'End If
		'myConn.Close()



		'Dim oDate As DateTime = DateTime.Parse(tglx(i))
		'tglsu = oDate.Day & " " & oDate.Month & "  " & oDate.Year
		'MsgBox(oDate.Day & " " & oDate.Month & "  " & oDate.Year)

		'Dim dt As DateTime = Convert.ToDateTime(tglxx)
		'tglsu = dt.Day & " " & dt.Month & "  " & dt.Year
		'tglsu = format(dt, "d-M-yyyy")
		'Dim dt As DateTime = format(tglefektifx(i), "yyyy-MM-dd")
		'tglefektifx(i) = dt
		'tglarr= tglx(i).Split("-")
		'tglx(i) = tglarr(2) & "-" & tglarr(1) & "-" & tglarr(0)
		
		datastrxx = idz(i) & ",'" & namaz(i) & "','" & nppz(i) & "'," & divzi(i) & "," & klpzi(i) & "," & unitzi(i) & "," & _
					posisizi(i) & "," & levelz(i) & "," & menusuratz(i) & "," & menusekz(i) & "," & menuadaz(i) & "," & menuasetz(i) & "," & _
					levelaksesx & "," & Val(posisiq)
						
		'datastrx(i) = idx(i) & "," & namax(i)
%>

<%
		If n=0 Then
			n = 1
%>
	<tr id="row<% Response.Write(bb)%>" STYLE="background-color: #FFFFFF" height="38px">
<%
		Elseif n=1 Then
			n = 0
%>
	<tr id="row<% Response.Write(bb)%>" STYLE="background-color: #E5FFFB" height="38px">
<%
		End If
%>
		<td style="padding-right: 10px; padding-top: 5px;  font-family: Arial; font-size: 0.9em;text-align: right; vertical-align: top; line-height: 110%"><% Response.Write (startval+bb) %></td>
		<td style="padding-left: 10px; padding-top: 5px;  font-family: Arial; font-size: 0.9em;text-align: left; vertical-align: top; line-height: 110%"><% Response.Write (namaz(i).ToUpper()) %></td>
		<td style="padding-left: 5px; padding-top: 5px;  font-family: Arial; font-size: 0.9em;text-align: center; vertical-align: top; line-height: 110%"><% Response.Write (nppz(i)) %></td>
		<td style="padding-left: 5px; padding-top: 5px;  font-family: Arial; font-size: 0.9em;text-align: center; vertical-align: top; line-height: 110%"><% Response.Write (divz(i)) %></td>
		<td style="padding-left: 5px; padding-top: 5px;  font-family: Arial; font-size: 0.9em;text-align: center; vertical-align: top; line-height: 110%"><% Response.Write (klpz(i)) %></td>
		<td style="padding-left: 5px; padding-top: 5px;  font-family: Arial; font-size: 0.9em;text-align: left; vertical-align: top; line-height: 110%"><% Response.Write (unitz(i)) %></td>
		<td style="padding-left: 5px; padding-top: 5px;  font-family: Arial; font-size: 0.9em;text-align: left; vertical-align: top; line-height: 110%"><% Response.Write (posisiz(i)) %></td>
		<!--td style="padding-left: 0px; font-family: Arial; font-size: 0.9em;text-align: center;">
			<a id="linkimage<% response.write(i) %>" href="<% Response.Write(namafilexx)%>" target="_blank">
			<img id="uploadPreview<% response.write(i) %>" src="<% Response.Write(namafilexx)%>" height="26px" width="25px"/><br/>
			</a>
		</td-->
		<td style="padding-left: 0px; padding-top: 5px; font-family: Arial; font-size: 0.9em; text-align: center; vertical-align: top;">
			<!--a id="edit<% Response.Write (i)%>" href="" class="small button blue" onclick="addkct(); return false;"><i class="fas fa-edit fa-sm"></i></a-->
			<a id="edit<% Response.Write (i)%>" href="#" onclick="editdatax(<% Response.Write(datastrxx) %>);">
			<img id="uploadPreview<% Response.Write (i)%>" src="./images/edit2.jpg" height="26px" width="25px"/><br/>
			</a>
		</td>
<%
		'If levelaksesx > 1 AND namaklpq.ToUpper() = "BUM" Then
		If levelaksesx > 1 Then
%>
		<td style="padding-left: 0px; padding-top: 5px; font-family: Arial; font-size: 0.9em;text-align: center; vertical-align: top;">
			<!--a id="hapus<% Response.Write(i)%>" href="#" class="small button red" onclick="addkct(); return false;"><i class="fas fa-trash fa-sm"></i></a-->		
			<img id="uploadPreview<% Response.Write (i)%>" src="./images/delete4.png" OnClick="hapususer(<% Response.Write (idz(i)) %>);" height="26px" width="25px"/><br/>
		</td>
<%
		End If
%>
	</tr>
<%
	next
	
Catch ex As Exception
        'lbl1.Text="Error while connecting Server."
		Response.Write (ex.Message)

Finally
		'myReader.Close()
        'myConn.Close() 'Whether there is error or not. Close the connection.	
End Try
%>

</table>

<table style="text-align: left; margin-left: 270px; margin-right: auto; margin-top: auto;" width="1140px" border="0" cellspacing="0" cellpadding="0">
	<tr height="30px">
		<td colspan="4" style="padding-left: 0px; text-align: left; font-size: 1.4em">

		</td>
	</tr>
</table>

<table id="totalpagetab" style="text-align: left; margin-left: 270px; margin-right: auto; margin-top: auto;" width="1140px" border="0" cellspacing="0" cellpadding="0">
	<tr height="30px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em">
			<FONT face="Arial" color="#000"><b>Total Hal:&nbsp;<% response.write(jmlpage) %></b></FONT>
		</td>
	</tr>
	<tr height="10px">
		<td style="padding-left: 0px; text-align: left; font-size: 1.4em">
		</td>
	</tr>
</table>

<table id="pagingtab" style="text-align: left; margin-left: 270px; margin-right: auto; margin-top: auto;" width="1100px" border="0" cellspacing="0" cellpadding="0">
	<tr height="34px" style="background-color: #FFFFFF">
		<td style="padding-left: auto; text-align: left; font-size: 1.1em;">
			<div class="pagination">
				<a href="#" onclick="viewpage(<% Response.Write(1) %>, <% Response.Write(jmlpage) %>)">&lt;&lt;</a>
				<a href="#" onclick="viewpage(<% Response.Write(pageno-1) %>, <% Response.Write(jmlpage) %>)">&lt;</a>
<%
			m = pageno \ 20
			n = pageno MOD 20
			startpage = (m * 20) + 1
			if n=0 AND m>0 Then
				startpage = ((m-1) * 20) + 1
			End If
			endpage = startpage + 19

			If endpage > jmlpage Then
				endpage = jmlpage
			End If
			
			for a=startpage to endpage
%>
			  <a 
			  <% If a=pageno Then 
					Response.Write ("class='active'")
				 End If
			  %>
				href="#" onclick="viewpage(<% Response.Write(a) %>, <% Response.Write(jmlpage) %>)"><% Response.Write(a) %></a>
<%
			next a
%>
				<a href="#" onclick="viewpage(<% Response.Write(pageno+1) %>, <% Response.Write(jmlpage) %>)">&gt;</a>
				<a href="#" onclick="viewpage(<% Response.Write(jmlpage) %>, <% Response.Write(jmlpage) %>)">&gt;&gt;</a>
			</div>
			
		</td>
		<td style="padding-left: 0px; text-align: left; font-size: 1.0em; vertical-align: middle;">
			
		</td>
	</tr>
	<tr height="30px">
		<td colspan="2" style="padding-left: 0px; text-align: left; font-size: 1.4em">
		</td>
	</tr>
</table>

<%
		Session("username") = user0
		Session("userid") = idq
		Session("namauser") = namauserq
		Session("idemployee") = nppq
		Session("levelakses") = levelaksesq
		Session("klp") = klpq
		Session("unit") = unitq
		Session("posisi") = posisiq
		Session("posisistr") = posisix
		Session("surat") = suratq
		Session("divisi") = divq
		Session("menusurat") = suratq
		Session("menusekre") = sekreq 
		Session("menuada") = adaq
		Session("menuaset") = asetq
		Session("statusaktif") = statusaktifq
%>


<form method="post" name="delfrm" id="delfrm" action="userdelgo.aspx" enctype="multipart/form-data">
	<input type="hidden" id="deluserid" name="deluserid">
</form>


</body>
</html>
