<!DOCTYPE html>

<%@ Page Language="VB" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data.Sql" %>
<%@ Import Namespace="System.Data.OleDb" %> 
<%@ Import Namespace="System.Configuration" %>
<%@ Import Namespace="System.Web.Configuration" %>

<%@ Import Namespace="ClosedXML.Excel" %>


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
<link rel="stylesheet" type="text/css" href="./tigracalendar/tcal.css" />
<script type="text/javascript" src="./tigracalendar/tcal.js"></script>

<link href="./datepicker4/dcalendar.picker.css" rel="stylesheet" type="text/css">

<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">



<style>
.example-element {
	margin-top: 0;
	position: fixed;
	background-color: #F9FBFB;
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
  padding: 17px 0px 17px 0px;
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
	color: #555555;
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
	color: #6681FF;
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
	color: #6D8FEB;
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
   width: 150px;
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
   padding: 2px 8px;
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
   padding: 2px 8px;
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
select#soflow2, select#soflow-color {
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
   width: 150px;
}

select#soflow2-color {
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
   padding: 2px 8px;
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
   padding: 2px 8px;
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
select#soflow23, select#soflow23-color {
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
   width: 150px;
}

select#soflow23-color {
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
label.custom-checkbox input[type="checkbox"] { 
    opacity:0;
}
 
label.custom-checkbox input[type="checkbox"] ~ .helping-el{
    background-color: #FFFFFF;
    border: 2px solid #009688;  
    border-radius: 2px;
    display: inline-block;
    padding: 10px;
    position: relative;
    top: 8px;
}
 
label.custom-checkbox input[type="checkbox"]:checked ~ .helping-el{ 
    border: 2px solid #009688;
    background-color: #009688;
}
 
label.custom-checkbox input[type="checkbox"]:checked ~ .helping-el:after {
    color: #FFFFFF;
    content: "\2714";
    font-size: 20px;
    font-weight: normal;
    left: 2px;
    position: absolute;
    top: -6px;
    transform: rotate(10deg);
}
</style>

<style>
label.custom-radio-button input[type="radio"] {
    opacity:0;
}
 
label.custom-radio-button input[type="radio"] ~ .helping-el {
    background-color: #FFFFFF;
    border: 2px solid #117CFF;  
    border-radius: 50%;
    display: inline-block;
    margin-right: 7px;
    padding: 11px;
    position: relative;
    top: 8px;
}
 
label.custom-radio-button input[type="radio"]:checked ~ .helping-el {
    border: 2px solid #0052B7;
}
 
label.custom-radio-button input[type="radio"]:checked ~ .helping-el:after {
    background-color: #117CFF;
    border-radius: 50%;
    content: " ";
    font-size: 30px;
    height: 16px;
    left: 3px;
    position: absolute;
    top: 3px;
    width: 16px;
     
}
</style>

<style>

#logoleft {
position: absolute;
top: 40px;
left: 40px;
}

</style>

<script>
function addkct()
{
	document.getElementById('kctbarutab').style.display = 'block';
	//document.getElementById('barukct').style.display = 'none';
}
</script>

<script>
function reset()
{
	document.getElementsByName("namamohon")[0].value = "";
	document.getElementsByName("tglperiode1")[0].value = "";
	document.getElementsByName("tglperiode2")[0].value = "";
	document.getElementsByName("statusfilter")[0].value = 0;
	document.forms["filterfrm"].submit();
}
</script>

<script>
function viewpage(nohal)
{
var myval = document.getElementsByName("namanpp")[0].value;
if (myval==null || myval=="")
{
	var halink = "main.aspx?page=" + nohal + "&cari=none";
}
else
{
	var halink = "main.aspx?page=" + nohal + "&cari=" + myval;
}
//window.open(halink);
window.location.href = halink;
	//document.getElementsByName("pageno")[0].value = nohal;
	//document.forms["pagefrm"].submit();
}
</script>

<script>
function viewfilter()
{
	document.getElementsByName("statusfilter")[0].value = 1;
	document.forms["filterfrm"].submit();
}
</script>

<script>
function databaru()
{
	document.getElementById("databarutab").style.display = "block";
	document.getElementById("dataedittab").style.display = "none";
	document.getElementById("filtertab").style.display = "none";
	document.getElementById("mybar").style.display = "none";
	document.getElementById("mylist").style.display = "none";
	document.getElementById("paging").style.display = "none";
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
	document.getElementById("mybar").style.display = "block";
	document.getElementById("mylist").style.display = "block";
	document.getElementById("paging").style.display = "block";
	//document.getElementById("clearbtn").style.display = "none";
	//document.getElementById("filtertab").style.display = "none";
}
</script>

<script>
function editdataxx(id,tahun,jenis,nama,tgl,tglx,jenisang,nilai,nostep,statusrev,folderx,filesur,fileip,filetor,filerab,posisi,alasan)
{
//alert(id);
	document.getElementsByName("padaid")[0].value = id;
	document.forms["myeditfrm"].submit();
}
</script>

<script>
function editdatax(id,tahun,jenis,nama,tgl,tglx,jenisang,nilai,nostep,statusrev,folderx,filesur,fileip,filetor,filerab,posisi,alasan)
{
	document.getElementById("databarutab").style.display = "none";
	document.getElementById("dataedittab").style.display = "block";
	document.getElementById("filtertab").style.display = "none";
	document.getElementById("mybar").style.display = "none";
	document.getElementById("mylist").style.display = "none";
	document.getElementById("paging").style.display = "none";
	document.getElementById("allbtn").style.display = "block";
	document.getElementById("batalbtn").style.display = "";

//alert(posisi + "--" + nostep + "++" + statusrev);
//alert(alasan);
	
	if (posisi < 4)			//asisten & analis
	{
		document.getElementById("aanbtn").style.display = "none";
		if (nostep==1)
		{
			document.getElementById("updatebtn").style.display = "";
			document.getElementById("approvebtn").style.display = "none";
			if (statusrev==1)
			{

				document.getElementById("revisibtn").style.display = "none";
			}
			else
			{
				document.getElementById("revisibtn").style.display = "none";
			}
		}
		else if (nostep>1)
		{
			document.getElementById("updatebtn").style.display = "none";
			document.getElementById("approvebtn").style.display = "none";
			document.getElementById("revisibtn").style.display = "none";
			document.getElementsByName("edittgl")[0].disabled = true;
			document.getElementsByName("editnama")[0].disabled = true;
			document.getElementById("soflow21").disabled = true;
			document.getElementById("soflow22").disabled = true;
			document.getElementById("soflow23").disabled = true;
			document.getElementsByName("editnom")[0].disabled = true;
			document.getElementById("edituploadImage1").disabled = true;
			document.getElementById("edituploadImage2").disabled = true;
			document.getElementById("edituploadImage3").disabled = true;
			document.getElementById("edituploadImage4").disabled = true;
			document.getElementById("alasantab").style.display = "block";
			document.getElementsByName("alasan")[0].disabled = true;
		}
		if (nostep>2)
		{
			document.getElementById("alasantab").style.display = "none";
		}
		if (nostep==5)
		{
			document.getElementById("awtab").style.display = "block";
			document.getElementById("aanbtn").style.display = "none";
		}
	}
	else if (posisi==4)		//Pemimpin Kelompok (4)
	{
		document.getElementsByName("edittgl")[0].disabled = true;
		document.getElementsByName("editnama")[0].disabled = true;
		document.getElementById("soflow21").disabled = true;
		document.getElementById("soflow22").disabled = true;
		document.getElementById("soflow23").disabled = true;
		document.getElementsByName("editnom")[0].disabled = true;
		document.getElementById("edituploadImage1").disabled = true;
		document.getElementById("edituploadImage2").disabled = true;
		document.getElementById("edituploadImage3").disabled = true;
		document.getElementById("edituploadImage4").disabled = true;
		document.getElementById("revisibtn").style.display = "none";
		document.getElementsByName("alasan")[0].disabled = true;
		//document.getElementById("alasan").readonly = true;
		if (nostep==1)
		{
			document.getElementById("updatebtn").style.display = "none";
			document.getElementById("approvebtn").style.display = "none";
		}
		else if (nostep==2)
		{
			document.getElementById("updatebtn").style.display = "none";
			document.getElementById("approvebtn").style.display = "display";
		}
		else if (nostep > 2)
		{
			document.getElementById("updatebtn").style.display = "none";
			document.getElementById("approvebtn").style.display = "none";
		}
	}
	else if (posisi==5)		//Pengelola/Manager (5)
	{
		document.getElementById("batalbtn").style.display = "block";
		if (nostep==1)
		{
			document.getElementById("updatebtn").style.display = "none";
			document.getElementById("approvebtn").style.display = "display";
		}
		else if (nostep>1)
		{
			document.getElementById("revisibtn").style.display = "none";
			document.getElementById("updatebtn").style.display = "none";
			document.getElementById("approvebtn").style.display = "none";
			document.getElementsByName("alasan")[0].disabled = true;
			document.getElementsByName("edittgl")[0].disabled = true;
			document.getElementsByName("editnama")[0].disabled = true;
			document.getElementById("soflow21").disabled = true;
			document.getElementById("soflow22").disabled = true;
			document.getElementById("soflow23").disabled = true;
			document.getElementsByName("editnom")[0].disabled = true;
			document.getElementById("edituploadImage1").disabled = true;
			document.getElementById("edituploadImage2").disabled = true;
			document.getElementById("edituploadImage3").disabled = true;
			document.getElementById("edituploadImage4").disabled = true;
		}
		if (statusrev==1)
		{
			document.getElementById("revisibtn").style.display = "none";
			document.getElementById("approvebtn").style.display = "none";
			if (nostep>1)
			{
				document.getElementById("alasantab").style.display = "block";
				document.getElementsByName("alasan")[0].disabled = false;
			}
		}
		if (nostep>2)
		{
			document.getElementById("alasantab").style.display = "none";
		}
	}
	else if (posisi==6 || posisi==11 || posisi==12 || posisi==13)		//DGM / Wakil Pemimpin (6)
	{
		document.getElementById("revisibtn").style.display = "none";
		document.getElementsByName("alasan")[0].disabled = true;
		document.getElementsByName("edittgl")[0].disabled = true;
		document.getElementsByName("editnama")[0].disabled = true;
		document.getElementById("soflow21").disabled = true;
		document.getElementById("soflow22").disabled = true;
		document.getElementById("soflow23").disabled = true;
		document.getElementsByName("editnom")[0].disabled = true;
		document.getElementById("edituploadImage1").disabled = true;
		document.getElementById("edituploadImage2").disabled = true;
		document.getElementById("edituploadImage3").disabled = true;
		document.getElementById("edituploadImage4").disabled = true;
		if (nostep < 3)
		{
			document.getElementById("updatebtn").style.display = "none";
			document.getElementById("approvebtn").style.display = "none";
		}
		else if (nostep==3)
		{
			document.getElementById("updatebtn").style.display = "none";
			document.getElementById("approvebtn").style.display = "display";
		}
		else if (nostep > 3)
		{
			document.getElementById("updatebtn").style.display = "none";
			document.getElementById("approvebtn").style.display = "none";
		}
	}
	else if (posisi==7)		//Pemimpin Divisi
	{
		document.getElementById("revisibtn").style.display = "none";
		document.getElementsByName("alasan")[0].disabled = true;
		document.getElementsByName("edittgl")[0].disabled = true;
		document.getElementsByName("editnama")[0].disabled = true;
		document.getElementById("soflow21").disabled = true;
		document.getElementById("soflow22").disabled = true;
		document.getElementById("soflow23").disabled = true;
		document.getElementsByName("editnom")[0].disabled = true;
		document.getElementById("edituploadImage1").disabled = true;
		document.getElementById("edituploadImage2").disabled = true;
		document.getElementById("edituploadImage3").disabled = true;
		document.getElementById("edituploadImage4").disabled = true;
		if (nostep < 4)
		{
			document.getElementById("updatebtn").style.display = "none";
			document.getElementById("approvebtn").style.display = "none";
		}
		else if (nostep==4)
		{
			document.getElementById("updatebtn").style.display = "none";
			document.getElementById("approvebtn").style.display = "display";
		}
		else if (nostep > 4)
		{
			document.getElementById("updatebtn").style.display = "none";
			document.getElementById("approvebtn").style.display = "none";
		}
	}

	document.getElementsByName("alasan")[0].value = alasan;
	document.getElementsByName("edittgl")[0].value = tglx;
	document.getElementsByName("edittahun")[0].value = tahun;
	document.getElementsByName("editnama")[0].value = nama;
	document.getElementsByName("editjenis")[0].value = jenis;
	document.getElementsByName("editjenisanggaran")[0].value = jenisang;
	document.getElementsByName("editnom")[0].value = nilai.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",");
	document.getElementsByName("enostep")[0].value = nostep;
	document.getElementsByName("eposisi")[0].value = posisi;
	document.getElementsByName("anostep")[0].value = nostep;
	document.getElementsByName("aposisi")[0].value = posisi;
	document.getElementsByName("aadaid")[0].value = id;
	document.getElementsByName("rnostep")[0].value = nostep;
	document.getElementsByName("rposisi")[0].value = posisi;
	document.getElementsByName("radaid")[0].value = id;
	document.getElementsByName("emohonid")[0].value = id;
	
	var fsur = "./images/nofile.png";
	if (filesur.length > 1)
	{
		fsur = "." + folderx + filesur;
	}
	document.getElementById("edituploadPreview1").src = fsur;
	document.getElementById("editlinkimage1").href = fsur;

	var resarr = filesur.split(".");
	var extens = resarr[1];
	if (extens=="pdf")
	{
		document.getElementById("edituploadPreview1").src = "./images/pdficon1.png";
	}
//-----------------------------
	var fip = "./images/nofile.png";
	if (fileip.length > 1)
	{
		fip = "." + folderx + fileip;
	}
	document.getElementById("edituploadPreview2").src = fip;
	document.getElementById("editlinkimage2").href = fip;

	resarr = fileip.split(".");
	extens = resarr[1];
	if (extens=="pdf")
	{
		document.getElementById("edituploadPreview2").src = "./images/pdficon1.png";
	}
//-----------------------------
	var ftor = "./images/nofile.png";
	if (filetor.length > 1)
	{
		ftor = "." + folderx + filetor;
	}
	document.getElementById("edituploadPreview3").src = ftor;
	document.getElementById("editlinkimage3").href = ftor;
	resarr = filetor.split(".");
	extens = resarr[1];
	if (extens=="pdf")
	{
		document.getElementById("edituploadPreview3").src = "./images/pdficon1.png";
	}
//-----------------------------
	var frab = "./images/nofile.png";
	if (filerab.length > 1)
	{
		frab = "." + folderx + filerab;
	}
	document.getElementById("edituploadPreview4").src = frab;
	document.getElementById("editlinkimage4").href = frab;
	resarr = filerab.split(".");
	extens = resarr[1];
	if (extens=="pdf")
	{
		document.getElementById("edituploadPreview4").src = "./images/pdficon1.png";
	}
}
</script>

<script>
function hapusada(id)
{
	if (confirm("Apakah anda ingin hapus data ?") == true) 
	{
		var tenure = prompt("Silahkan masukkan kode hapus", "");
		
		if (tenure == "123") {
			document.getElementsByName("deladaid")[0].value = id;
			document.forms["delfrm"].submit();
		}
	}
}
</script>

<script>
function approve()
{
	document.forms["approvefrm"].submit();
}
</script>

<script>
function savebaru()
{
	document.forms["barufrm"].submit();
}
</script>

<script>
function saveedit()
{
	document.forms["editfrm"].submit();
}
</script>

<script>
function saverevisi()
{
	var alasan = document.getElementsByName("alasan")[0].value;
	var alasanlen = alasan.length;
	if (alasanlen<2)
	{
		alert("Masukkan alasan revisi");
		return false;
	}
	else
	{
		document.getElementsByName("ralasan")[0].value = alasan
		document.forms["revisifrm"].submit();
	}
}
</script>

<script type="text/javascript">     
function PreviewImage(no) 
{         
	var oFReader = new FileReader();
	var namafil = document.getElementById("uploadImage"+no).value;
	var namafilearr = namafil.split(".");
	var noarr = namafilearr.length;
	var jenisfile = namafilearr[noarr-1];

	oFReader.readAsDataURL(document.getElementById("uploadImage"+no).files[0]);
	
	if (jenisfile==='pdf')
	{
		oFReader.onload = function (oFREvent) {             
		document.getElementById("uploadPreview"+no).src = './images/pdficon1.png'; 	
		};
	}
	else
	{
		oFReader.onload = function (oFREvent) {             
		document.getElementById("uploadPreview"+no).src = oFREvent.target.result; 		
		};
	}
	//document.forms["memobarufrm"].submit();
	//memousulanbaru();

} 
</script>

<script type="text/javascript">     
function editPreviewImage(no) 
{         
	var oFReader = new FileReader();
	var namafil = document.getElementById("edituploadImage"+no).value;
	var namafilearr = namafil.split(".");
	var noarr = namafilearr.length;
	var jenisfile = namafilearr[noarr-1];

	oFReader.readAsDataURL(document.getElementById("edituploadImage"+no).files[0]);
	
	if (jenisfile==='pdf')
	{
		oFReader.onload = function (oFREvent) {             
		document.getElementById("edituploadPreview"+no).src = './images/pdficon1.png'; 	
		};
	}
	else
	{
		oFReader.onload = function (oFREvent) {             
		document.getElementById("edituploadPreview"+no).src = oFREvent.target.result; 		
		};
	}
	//document.forms["memobarufrm"].submit();
	//memousulanbaru();

} 
</script>

<script>
function formatnumber(nField) {
//alert(this);
  //el.value = el.value.replace(/[^\d]/g,'').replace(/(\d\d?)$/,'.$1');
		if (/^0/.test(nField.value))
			{
			 nField.value = nField.value.substring(0,1);
			}
		if (Number(nField.value.replace(/,/g,"")))
			{
			 var tmp = nField.value.replace(/,/g,"");
	 		 tmp = tmp.toString().split('').reverse().join('').replace(/(\d{3})/g,'$1,').split('').reverse().join('').replace(/^,/,'');
			 if (/\./g.test(tmp))
				{
			 	 tmp = tmp.split(".");
				 tmp[1] = tmp[1].replace(/\,/g,"").replace(/ /,"");
				 nField.value = tmp[0]+"."+tmp[1]
				}
			 else 	{
				 nField.value = tmp.replace(/ /,"");
				} 
			}
		else	{
			 nField.value = nField.value.replace(/[^\d\,\.]/g,"").replace(/ /,"");
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
	Dim a, aaa, b, bb, c, d, jmla, jmldata as integer
	Dim user0 as string
	Dim pwd0 as string
	Dim namauserx as String
	Dim aksesx as String
	Dim namamohon as String = ""
	Dim idx as String
	Dim nosuratx as String
	Dim tglx as String
	Dim tahunx as String
	Dim jenisx as String
	Dim jenisangx as String = ""
	Dim namax as String
	Dim nilaix as Long = 0
	Dim tembusanx as String
	Dim cuserx as String
	Dim nostepx as String
	Dim nostepijinx as String
	Dim stepx as String
	Dim statusrevisix as String
	Dim filenamex as String
	Dim statusf as Integer = 0
	'Dim pageno as Integer = 1
	
	Dim rowperpage as Integer = 20
	Dim pageno as String
	Dim startpage as Long = 0
	Dim endpage as Long = 0
	Dim jmlpage as Integer = 1
	Dim pagerows as Integer = 20
	Dim rowsperpage as Integer = 20
	Dim hal as Integer = 1
	
	Dim useridx as Integer = 0
	Dim itop as Integer = 1
	Dim ii as Integer = 0
	Dim k as Integer
	Dim userid as Integer = 1
	
	Dim idz(1000) as String
	Dim idq As String = ""
	Dim namauserq As String = ""
	Dim statusq As String = ""
	Dim jabatanq As String = ""
	Dim klpq As String = ""
	Dim unitq As String = ""
	Dim posisiq As String = ""
	Dim posisint as Integer = 0
	Dim suratq As String = ""
	Dim sekreq As String = ""
	Dim divq As String = ""
	Dim nppq as String = ""
	Dim levelaksesq as String = ""
	Dim adaq as String = ""
	Dim asetq as String = ""
	Dim statusaktifq as String = ""
	Dim posisix as String = ""
	Dim namastr as String = ""
	Dim tglstr as String = ""
	Dim tgl1 as String = ""
	Dim tgl2 as String = ""
	Dim tglqu as String = ""
	Dim tahap as String  = ""
	Dim tahapstr as String = ""
	Dim loguserstr as String = ""
	Dim namaklpq as String = ""


	Dim SQL0 As String
	Dim SQL1 As String
	Dim SQL2 As String
	Dim SQL3 as String
	Dim nama20 as String = ""
	Dim posisi20 as String = ""
	Dim unitstr as String = ""
	Dim klpstr as String = ""
	Dim divstr as String = ""
	
	Dim perihal as String = ""
	Dim perihalstr as String = ""
	
	Dim startval as Integer = 0
	Dim endval as Integer = 0
	Dim datastrx as String
	Dim tahuniki as Integer
	Dim tahunwingi as Integer
	Dim tahunsesuk as Integer
	Dim tgl10 as String = ""
	Dim tgl20 as String = ""
	Dim tglarr(10) as String
	Dim stepstr as String = ""
	Dim alasanx as String = ""
	
	Dim setmenu as String = "mohon"

	Dim regDate as Date = Date.Now()
	Dim strDate as String = regDate.ToString("dd-MM-yyyy")

	'lbl1.Text = Hour(Now) & ":" & Minute(Now) & ":" & Second(Now)
	
	tahuniki = Val(Year(Now))
	tahunwingi = tahuniki - 1
	tahunsesuk = tahuniki + 1

	aaa=1
	
	user0 = ""
	pwd0 = ""
	

    Try
        If aaa=1 Then
			'myConn = New SqlConnection("Data Source=DESKTOP-S7N8DF9;Initial Catalog=KCT;User ID=sa;Password=sa;")
%>

<!--#include file="koneksi.aspx"-->

<%
			If Request.Form("pageno") isNot Nothing Then
				pageno = Request.Form("pageno")
				hal = Cint(pageno)
			End If
			If NOT Request.Form("namamohon") is Nothing Then
				namamohon = Request.Form("namamohon")
			End If
			If NOT Request.Form("tglperiode1") is Nothing Then
				tgl1 = Request.Form("tglperiode1")
			End If
			If NOT Request.Form("tglperiode2") is Nothing Then
				tgl2 = Request.Form("tglperiode2")
			End If
			If NOT Request.Form("tahapan") is Nothing Then
				tahap = Request.Form("tahapan")
			End If
			If NOT Request.Form("statusfilter") is Nothing Then
				statusf = Cint(Request.Form("statusfilter"))
			End If
			
			user0 = Session("username")
			idq = Session("userid")
			nppq = Session("idemployee")
			If not Session("namauser") is nothing Then
				namauserq = Session("namauser")
				nama20 = Left(namauserq, 16)
			End If
			jabatanq = Session("jabatan")
			klpq = Session("klp")
			If not Session("namaklp") is nothing Then
				namaklpq = Session("namaklp")
			End If
			unitq = Session("unit")
			posisiq = Session("posisi")
			posisint = Cint(posisiq)
			posisix = Session("posisistr")
			suratq = Session("menusurat")
			sekreq = Session("menusekre")
			adaq = Session("menuada")
			asetq = Session("menuaset")
			divq = Session("divisi")
			levelaksesq = Session("levelakses")
			posisi20 = Left(posisix, 20)

			If userid=0 Then
				userid = idq
			End If
			
			If namauserq.length()<2 Then
%>

<script>
alert('Anda tidak beraktivitas terlalu lama (idle). Silahkan signin lagi');
window.top.location.href = "index.html"; 
</script>

<%			
			End If
			
			
			myCmd = myConn.CreateCommand

'** Baca semua data sesuai Query
			namastr = ""
			If namamohon.length()>0 Then
				namastr = " AND (nama LIKE '%" & namamohon & "%')"
			End If
			
			tahapstr = ""
			If Val(tahap) = 0 Then
				tahapstr = ""
			Elseif Val(tahap) > 1 Then
				tahapstr = " AND (nostep=" & tahap & ")"
			End If
			
			If klpq.length()>0 Then
				klpstr = " AND klp=" & klpq
			End If
			If divq.length()>0 Then
				divstr = " AND divisi=" & divq
			End If
			
			tglstr = ""
			If tgl1.length() > 1 AND String.IsNullOrEmpty(tgl2) Then
				tglarr = tgl1.Split("-")
				tgl10 = tglarr(2) & "-" & tglarr(1) & "-" & tglarr(0)
				tglstr = " AND tglmohon='" & tgl10 & "'"
			Elseif String.IsNullOrEmpty(tgl1) AND tgl2.length() > 1 Then
				tglarr = tgl2.Split("-")
				tgl20 = tglarr(2) & "-" & tglarr(1) & "-" & tglarr(0)
				tglstr = " AND tglmohon='" & tgl20 & "'"
			Elseif tgl1.length()>1 AND tgl2.length()>1 Then
				tglarr = tgl1.Split("-")
				tgl10 = tglarr(2) & "-" & tglarr(1) & "-" & tglarr(0)
				tglarr = tgl2.Split("-")
				tgl20 = tglarr(2) & "-" & tglarr(1) & "-" & tglarr(0)
				tglstr = " AND tglmohon BETWEEN '" & tgl10 & "' AND '" & tgl20 & "'"
			End If	

			If posisint<4 Then
				loguserstr = " AND (klp=" & klpq & ")"
			ElseIf posisint=4 Then
				loguserstr = " AND (klp=" & klpq & ")"
			ElseIf posisint=5 Then
				loguserstr = " AND (klp=" & klpq & ")"
			ElseIf posisint>5 Then
				loguserstr = " AND (div=" & divq & ")"
			End If
			
        End If

    Catch ex As Exception
        'lbl1.Text="Error while connecting Server."
		Response.Write (ex.Message)

    Finally
		'myReader.Close()
        'myConn.Close() 'Whether there is error or not. Close the connection.

    End Try
%>
<div class="example-element"></div>
 
<!--#include file="menu.aspx"-->

<!--#include file="topbar.aspx"-->

<div class="container">
	<!--div class="logo"><img id = "logoleft" src= "./images/portalbumlogo_red.png" style="width:180px;height:90px;"/></div-->
	<div class="welcome">Selamat Datang,</div>
	<div class="photo"><img src="./images/user5.png" alt="" style="width:40px;height:40px;"></div>
	<div class="userin"><i><b><% Response.Write(nama20) %></b></i></div>
	<div class="posisi" STYLE="font-size: 0.8em"><i><% Response.Write(posisi20) %></i></div>
</div>

<table style="text-align: left; margin-left: 270px; margin-right: auto; margin-top: auto;" width="950px" border="0" cellspacing="0" cellpadding="0">
	<tr height="30px">
		<td colspan="4" style="padding-left: 20px; text-align: left; font-size: 1.0em">
		</td>
	</tr>
	<tr height="20px">
		<td colspan="4" style="padding-left: 0px; text-align: left; font-size: 1.5em; font-weight: 500">
			<FONT face="Arial" color="#000">PERMOHONAN PENGADAAN</FONT>
		</td>
	</tr>
	<tr height="30px">
		<td colspan="4" style="padding-left: 20px; text-align: left; font-size: 1.0em">
		
		</td>
	</tr>
</table>

<%
'response.write("-------------------------------------------------" + posisiq)
%>

<form method="post" name="barufrm" id="barufrm" action="permohonanaddgo.aspx" enctype="multipart/form-data">
<table id="databarutab" style="text-align: left; margin-left: 270px; margin-right: auto; margin-top: 0; display: none" width="990px" border="0" cellspacing="0" cellpadding="0">
	<tr height="20px">
		<td colspan="4" style="padding-left: 0px; text-align: left; font-size: 1.3em">
			<FONT face="Arial" color="#44AA44">DATA BARU</FONT>
		</td>
	</tr>
	<tr height="40px">
		<td colspan="4" style="padding-left: 20px; text-align: left; font-size: 1.0em">
		</td>
	</tr>
	<tr height="35px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top" width="330px">
			<FONT face="Arial" color="#000">
				Tgl 
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top" width="30px">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: top" width="550px">
			<FONT face="Arial" color="#000">
				<!--input style="font-family: Arial; font-size: 0.9em" type="text" id="datepicker-ex30" name="editbbmtglbayar" class="textboxtgl"-->
				<input autocomplete="off" STYLE="font-family: Arial; padding-left:5px; font-size:0.9em; border:1px solid #ababab; height: 24px; width: 120px" autocomplete="off" class="tcal" id="barutgl" name="barutgl" type="text" tabindex="1">
			</FONT>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em" width="80px">										
		</td>
	</tr>
	<tr height="35px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Tahun Pengadaan
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; font-size: 1.0em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: top">
			<FONT face="Arial" color="#000">
			<select STYLE="color: #000000; background-color: #ffffff; font-size: 0.9em; width: 100px" id="soflow11" name="barutahun" data-placeholder="Jenis" tabindex="2"> 
				<option value="" >&nbsp;&nbsp;</option>
				<option value="<% Response.Write(tahunwingi)%>"><% Response.Write(tahunwingi)%></option>
				<option value="<% Response.Write(tahuniki)%>" selected><% Response.Write(tahuniki)%></option>
				<option value="<% Response.Write(tahunsesuk)%>"><% Response.Write(tahunsesuk)%></option>
			</select>			
			</FONT>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em">										
		</td>
	</tr>
	<tr height="65px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Nama Pengadaan
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; font-size: 1.0em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: top">
			<FONT face="Arial" color="#000">
				<textarea STYLE="font-family: Arial; color: #000000; padding-left: 5px; border:1px solid #ababab; font-size:1.0em; line-height: 110%; height: 50px; width: 680px;" id="barunama" name="barunama" tabindex="3"></textarea>
			</FONT>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em">										
		</td>
	</tr>
	<tr height="35px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Jenis Pengadaan
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; font-size: 1.0em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: top">
			<FONT face="Arial" color="#000">
			<select STYLE="color: #000000; background-color: #ffffff; font-size: 0.9em; width: 150px" id="soflow12" name="barujenis" data-placeholder="Jenis" tabindex="4"> 
				<option value="" >&nbsp;&nbsp;</option>
				<option value="SENTRALISASI" >SENTRALISASI</option>
				<option value="DESENTRALISASI" >DESENTRALISASI</option>
			</select>			
			</FONT>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em">										
		</td>
	</tr>
	<tr height="35px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Jenis Anggaran
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; font-size: 1.0em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: top">
			<FONT face="Arial" color="#000">
			<select STYLE="color: #000000; background-color: #ffffff; font-size: 0.9em; width: 150px" id="soflow13" name="barujenisanggaran" data-placeholder="Jenis" tabindex="5"> 
				<option value="" >&nbsp;&nbsp;</option>
				<option value="OPEX" >OPEX</option>
				<option value="CAPEX" >CAPEX</option>
			</select>			
			</FONT>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em">										
		</td>
	</tr>
	<tr height="35px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Nilai Anggaran
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: top">
			<FONT face="Arial" color="#000">
				<input style="font-family: Arial; padding-right: 10px; text-align: right; font-size: 1.0em; width: 150px; height: 20px" type="text" id="barunom" name="barunom" class="textbox" onkeyup="formatnumber(this);" tabindex="6">
			</FONT>
			<FONT face="Arial" color="#000" size="3">
			(Rp.)
			</FONT>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em">										
		</td>
	</tr>
	<!--tr height="35px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Nilai HPS
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: top">
			<FONT face="Arial" color="#000">
				<input style="padding-right: 10px; text-align: right; font-size: 1.0em; width: 150px; height: 20px" type="text" id="baruhps" name="baruhps" class="textbox" onkeyup="formatnumber(this);" tabindex="5">
			</FONT>
			<FONT face="Arial" color="#000" size="3">
			(Rp.)
			</FONT>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em">										
		</td>
	</tr-->

	<tr height="34px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Upload Berkas/Dok
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
			</FONT>											
		</td>
		<td style="padding-left: 0px; padding-top: 5px; text-align: left; vertical-align: top">
		</td>
	</tr>
	<tr height="30px">
		<td style="padding-left: 30px; padding-top:5px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Dok 1
			</FONT>											
		</td>
		<td style="padding-left: 0px; padding-top:5px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; padding-top: 5px; text-align: left; vertical-align: top">
			<table valign="top" style="text-align: left; margin-left: 0px; margin-right: auto; margin-top: 0px;" width="600px" border="0" cellspacing="0" cellpadding="0">
				<tr height="42px">
					<td style="padding-left: 0px; text-align: left; vertical-align: top" width="400px">
						<FONT face="Arial" color="#000">
							<input style="padding-left: 0px; text-align: left; font-size: 0.9em" id="uploadImage1" name="uploadImage1" type="file" onchange="PreviewImage(1);" />
						</FONT>
					</td>
					<td style="padding-left: 0px; text-align: left; vertical-align: top" width="200px">
						<FONT face="Arial" color="#000">
							<a id="linkimage1" href="#" target="_blank">
							<img id="uploadPreview1" src="./images/nofile.png" height="40px" width="30px"/><br/>
							</a>
						</FONT>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr height="30px">
		<td style="padding-left: 30px; padding-top:5px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Dok 2
			</FONT>											
		</td>
		<td style="padding-left: 0px; padding-top:5px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; padding-top: 5px; text-align: left; vertical-align: top">
			<table id="file1" style="text-align: left; margin-left: 0px; margin-right: auto; margin-top: 0px;" width="600px" border="0" cellspacing="0" cellpadding="0">
				<tr height="42px">
					<td style="padding-left: 0px; text-align: left; vertical-align: top" width="400px">
						<FONT face="Arial" color="#000">
							<input style="padding-left: 0px; text-align: left; font-size: 0.9em" id="uploadImage2" name="uploadImage2" type="file" onchange="PreviewImage(2);" />
						</FONT>
					</td>
					<td style="padding-left: 0px; text-align: left; vertical-align: top" width="200px">
						<FONT face="Arial" color="#000">
							<a id="linkimage2" href="#" target="_blank">
							<img id="uploadPreview2" src="./images/nofile.png" height="40px" width="30px"/><br/>
							</a>
						</FONT>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr height="30px">
		<td style="padding-left: 30px; padding-top:5px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Dok 3
			</FONT>											
		</td>
		<td style="padding-left: 0px; padding-top:5px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; padding-top: 5px; text-align: left; vertical-align: top">
			<table id="file1" style="text-align: left; margin-left: 0px; margin-right: auto; margin-top: 0px;" width="600px" border="0" cellspacing="0" cellpadding="0">
				<tr height="42px">
					<td style="padding-left: 0px; text-align: left; vertical-align: top" width="400px">
						<FONT face="Arial" color="#000">
							<input style="padding-left: 0px; text-align: left; font-size: 0.9em" id="uploadImage3" name="uploadImage3" type="file" onchange="PreviewImage(3);" />
						</FONT>
					</td>
					<td style="padding-left: 0px; text-align: left; vertical-align: top" width="200px">
						<FONT face="Arial" color="#000">
							<a id="linkimage3" href="#" target="_blank">
							<img id="uploadPreview3" src="./images/nofile.png" height="40px" width="30px"/><br/>
							</a>
						</FONT>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr height="30px">
		<td style="padding-left: 30px; padding-top:5px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Dok 4
			</FONT>											
		</td>
		<td style="padding-left: 0px; padding-top:5px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; padding-top: 5px; text-align: left; vertical-align: top">
			<table id="file1" style="text-align: left; margin-left: 0px; margin-right: auto; margin-top: 0px;" width="600px" border="0" cellspacing="0" cellpadding="0">
				<tr height="42px">
					<td style="padding-left: 0px; text-align: left; vertical-align: top" width="400px">
						<FONT face="Arial" color="#000">
							<input style="padding-left: 0px; text-align: left; font-size: 0.9em" id="uploadImage4" name="uploadImage4" type="file" onchange="PreviewImage(4);" />
						</FONT>
					</td>
					<td style="padding-left: 0px; text-align: left; vertical-align: top" width="200px">
						<FONT face="Arial" color="#000">
							<a id="linkimage4" href="#" target="_blank">
							<img id="uploadPreview4" src="./images/nofile.png" height="40px" width="30px"/><br/>
							</a>
						</FONT>
					</td>
				</tr>
			</table>
		</td>
	</tr>

	<tr height="20px">
		<td colspan="3" style="padding-left: 20px; text-align: left;">
		</td>
	</tr>
	<tr height="40px">
		<td style="padding-left: 0px; text-align: left;">
			<a id="batalbtn" href="#" STYLE="font-family: Calibri; font-size: 1.0em;" class="medium button orange" onclick="hidedatabaru();"><i class="fa fa-close"></i>&nbsp;&nbsp;Tutup</a>
		</td>
		<td colspan="3" style="padding-left: 0px; text-align: left;">
			<a id="simpanbtn" href="#" STYLE="font-family: Calibri; font-size: 1.0em;" class="medium button green" onclick="savebaru();"><i class="fa fa-save" aria-hidden="true"></i>&nbsp;&nbsp;Ajukan</a>
		</td>
	</tr>
	<tr height="30px">
		<td colspan="4">
			<input type="hidden" id="kctid" name="kctid">
			<input type="hidden" id="editfile" name="editfile">
			<input type="hidden" id="baruhps" name="baruhps" value="0">
		</td>
	</tr>
</table>
</form>

<%
'response.write("-----------------------------------" & posisiq & "--" & posisix)
'====================================================================================================
%>

<form method="post" name="editfrm" id="editfrm" action="pengadaaneditgo.aspx" enctype="multipart/form-data">
<table id="dataedittab" style="text-align: left; margin-left: 270px; margin-right: auto; margin-top: 0; display: none" width="880px" border="0" cellspacing="0" cellpadding="0">
	<tr height="20px">
		<td colspan="4" style="padding-left: 0px; text-align: left; font-size: 1.3em">
			<FONT face="Arial" color="#44AA44"><b>REVIEW/UPDATE</b></FONT>
		</td>
	</tr>
	<tr height="40px">
		<td colspan="4" style="padding-left: 20px; text-align: left; font-size: 1.0em">
		</td>
	</tr>
	<tr height="35px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top" width="220px">
			<FONT face="Arial" color="#000">
				Tgl 
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 0.9em; vertical-align: top" width="30px">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: top" width="550px">
			<FONT face="Arial" color="#000">
				<!--input style="font-family: Arial; font-size: 0.9em" type="text" id="datepicker-ex30" name="editbbmtglbayar" class="textboxtgl"-->
				<input autocomplete="off" STYLE="padding-left:5px; font-size:0.9em; border:1px solid #ababab; height: 24px; width: 120px" autocomplete="off" class="tcal" id="edittgl" name="edittgl" type="text" tabindex="1">
			</FONT>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em" width="80px">										
		</td>
	</tr>
	<tr height="35px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Tahun Pengadaan
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.0em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: top">
			<FONT face="Arial" color="#000">
			<select STYLE="color: #000000; background-color: #ffffff; font-size: 0.9em; width: 150px" id="soflow21" name="edittahun" data-placeholder="Jenis" tabindex="2"> 
				<option value="" >&nbsp;&nbsp;</option>
				<option value="<% Response.Write(tahunwingi)%>"><% Response.Write(tahunwingi)%></option>
				<option value="<% Response.Write(tahuniki)%>"><% Response.Write(tahuniki)%></option>
				<option value="<% Response.Write(tahunsesuk)%>"><% Response.Write(tahunsesuk)%></option>
			</select>			
			</FONT>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em">										
		</td>
	</tr>
	<tr height="35px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Jenis Pengadaan
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.0em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: top">
			<FONT face="Arial" color="#000">
			<select STYLE="color: #000000; background-color: #ffffff; font-size: 0.9em; width: 150px" id="soflow22" name="editjenis" data-placeholder="Jenis" tabindex="2"> 
				<option value="" >&nbsp;&nbsp;</option>
				<option value="Sentralisasi" >Sentralisasi</option>
				<option value="Desentralisasi" >Desentralisasi</option>
			</select>			
			</FONT>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em">										
		</td>
	</tr>
	<tr height="65px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Nama Pengadaan
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.0em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: top">
			<FONT face="Arial" color="#000">
				<textarea STYLE="color: #000000; font-family: Arial; padding-left: 5px; border:1px solid #ababab; font-weight: normal; font-size:1.0em; line-height: 110%;" id="editnama" name="editnama" rows="3" cols="81" tabindex="3"></textarea>
			</FONT>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em">										
		</td>
	</tr>
	<tr height="35px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Jenis Anggaran
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.0em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: top">
			<FONT face="Arial" color="#000">
			<select STYLE="color: #000000; background-color: #ffffff; font-size: 0.9em; width: 150px" id="soflow23" name="editjenisanggaran" data-placeholder="Jenis" tabindex="2"> 
				<option value="" >&nbsp;&nbsp;</option>
				<option value="opex" >OPEX</option>
				<option value="capex" >CAPEX</option>
			</select>			
			</FONT>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em">										
		</td>
	</tr>
	<tr height="35px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Jumlah Anggaran
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: top">
			<FONT face="Arial" color="#000">
				<input style="font-family: Arial; font-size: 1.0em; width: 200px" type="text" id="editnom" name="editnom" class="textbox" onkeyup="formatnumber(this);" tabindex="4">
			</FONT>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em">										
		</td>
	</tr>
	<tr height="34px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Upload Berkas/Dok
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
			</FONT>											
		</td>
		<td colspan="2" style="padding-left: 0px; padding-top: 5px; text-align: left; vertical-align: top">
		</td>
	</tr>
	<tr height="30px">
		<td style="padding-left: 30px; padding-top:5px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Dok 1
			</FONT>											
		</td>
		<td style="padding-left: 0px; padding-top:5px; text-align: center; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td colspan="2" style="padding-left: 0px; padding-top: 5px; text-align: left; vertical-align: top">
			<table valign="top" style="text-align: left; margin-left: 0px; margin-right: auto; margin-top: 0px;" width="600px" border="0" cellspacing="0" cellpadding="0">
				<tr height="42px">
					<td style="padding-left: 0px; text-align: left; vertical-align: top" width="400px">
						<FONT face="Arial" color="#000">
							<input style="padding-left: 0px; text-align: left; font-size: 0.9em" id="edituploadImage1" name="edituploadImage1" type="file" onchange="editPreviewImage(1);" />
						</FONT>
					</td>
					<td style="padding-left: 0px; text-align: left; vertical-align: top" width="200px">
						<FONT face="Arial" color="#000">
							<a id="editlinkimage1" href="#" target="_blank">
							<img id="edituploadPreview1" height="40px" width="30px"/><br/>
							</a>
						</FONT>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr height="30px">
		<td style="padding-left: 30px; padding-top:5px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Dok 2
			</FONT>											
		</td>
		<td style="padding-left: 0px; padding-top:5px; text-align: center; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td colspan="2" style="padding-left: 0px; padding-top: 5px; text-align: left; vertical-align: top">
			<table id="file1" style="text-align: left; margin-left: 0px; margin-right: auto; margin-top: 0px;" width="600px" border="0" cellspacing="0" cellpadding="0">
				<tr height="42px">
					<td style="padding-left: 0px; text-align: left; vertical-align: top" width="400px">
						<FONT face="Arial" color="#000">
							<input style="padding-left: 0px; text-align: left; font-size: 0.9em" id="edituploadImage2" name="edituploadImage2" type="file" onchange="editPreviewImage(2);" />
						</FONT>
					</td>
					<td style="padding-left: 0px; text-align: left; vertical-align: top" width="200px">
						<FONT face="Arial" color="#000">
							<a id="editlinkimage2" href="#" target="_blank">
							<img id="edituploadPreview2" height="40px" width="30px"/><br/>
							</a>
						</FONT>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr height="30px">
		<td style="padding-left: 30px; padding-top:5px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Dok 3
			</FONT>											
		</td>
		<td style="padding-left: 0px; padding-top:5px; text-align: center; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td colspan="2" style="padding-left: 0px; padding-top: 5px; text-align: left; vertical-align: top">
			<table id="file1" style="text-align: left; margin-left: 0px; margin-right: auto; margin-top: 0px;" width="600px" border="0" cellspacing="0" cellpadding="0">
				<tr height="42px">
					<td style="padding-left: 0px; text-align: left; vertical-align: top" width="400px">
						<FONT face="Arial" color="#000">
							<input style="padding-left: 0px; text-align: left; font-size: 0.9em" id="edituploadImage3" name="edituploadImage3" type="file" onchange="editPreviewImage(3);" />
						</FONT>
					</td>
					<td style="padding-left: 0px; text-align: left; vertical-align: top" width="200px">
						<FONT face="Arial" color="#000">
							<a id="editlinkimage3" href="#" target="_blank">
							<img id="edituploadPreview3" height="40px" width="30px"/><br/>
							</a>
						</FONT>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr height="30px">
		<td style="padding-left: 30px; padding-top:5px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Dok 4
			</FONT>											
		</td>
		<td style="padding-left: 0px; padding-top:5px; text-align: center; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td colspan="2" style="padding-left: 0px; padding-top: 5px; text-align: left; vertical-align: top">
			<table id="file1" style="text-align: left; margin-left: 0px; margin-right: auto; margin-top: 0px;" width="600px" border="0" cellspacing="0" cellpadding="0">
				<tr height="42px">
					<td style="padding-left: 0px; text-align: left; vertical-align: top" width="400px">
						<FONT face="Arial" color="#000">
							<input style="padding-left: 0px; text-align: left; font-size: 0.9em" id="edituploadImage4" name="edituploadImage4" type="file" onchange="editPreviewImage(4);" />
						</FONT>
					</td>
					<td style="padding-left: 0px; text-align: left; vertical-align: top" width="200px">
						<FONT face="Arial" color="#000">
							<a id="editlinkimage4" href="#" target="_blank">
							<img id="edituploadPreview4" height="40px" width="30px"/><br/>
							</a>
						</FONT>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr height="35px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Status
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: top">
			<FONT face="Arial" color="#000">
			</FONT>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em">										
		</td>
	</tr>
</table>

<table id="alasantab" style="text-align: left; margin-left: 270px; margin-right: auto; margin-top: 0; display: none" width="880px" border="0" cellspacing="0" cellpadding="0">
	<tr height="65px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Alasan Revisi (jika ada)
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.0em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: top">
			<FONT face="Arial" color="#000">
				<textarea STYLE="color: #000000; font-family: Arial; padding-left: 5px; border:1px solid #ababab; font-weight: normal; font-size:1.0em; line-height: 110%;" id="alasan" name="alasan" rows="3" cols="81" tabindex="3"></textarea>
			</FONT>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em">										
		</td>
	</tr>
</table>
<table id="awtab" style="text-align: left; margin-left: 270px; margin-right: auto; margin-top: 0; display: none" width="880px" border="0" cellspacing="0" cellpadding="0">
	<tr height="35px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top" width="220px">
			<FONT face="Arial" color="#000">
				Tgl Aanwijzing
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 0.9em; vertical-align: top" width="30px">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: top" width="550px">
			<FONT face="Arial" color="#000">
				<!--input style="font-family: Arial; font-size: 0.9em" type="text" id="datepicker-ex30" name="editbbmtglbayar" class="textboxtgl"-->
				<input autocomplete="off" STYLE="padding-left:5px; font-size:0.9em; border:1px solid #ababab; height: 24px; width: 120px" autocomplete="off" class="tcal" id="edittglaw" name="edittglaw" type="text" tabindex="1">
			</FONT>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em" width="80px">										
		</td>
	</tr>
	<tr height="65px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Keterangan
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.0em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: top">
			<FONT face="Arial" color="#000">
				<textarea STYLE="color: #000000; font-family: Arial; padding-left: 5px; border:1px solid #ababab; font-weight: normal; font-size:1.0em; line-height: 110%;" id="alasan" name="alasan" rows="3" cols="81" tabindex="3"></textarea>
			</FONT>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em">										
		</td>
	</tr>
</table>
<table id="allbtn" style="text-align: left; margin-left: 270px; margin-right: auto; margin-top: 0; display: none" width="880px" border="0" cellspacing="0" cellpadding="0">
	<tr height="50px">
		<td colspan="4" style="padding-left: 20px; text-align: left;">
		</td>
	</tr>
	<tr height="40px">
		<td style="padding-left: 0px; text-align: left;" width="150px">
			<a id="batalbtn" href="#" STYLE="font-family: Calibri; font-size: 1.0em;" class="medium button orange" onclick="hidedataedit();">Tutup</a>
		</td>
		<td colspan="3" style="padding-left: 0px; text-align: left;" width="730px">
			<table style="text-align: left; margin-left: 10px; margin-right: auto; margin-top: 0px;" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td style="padding-left: 0px; text-align: left;" width="150px">
						<a id="revisibtn" href="#" STYLE="font-family: Calibri; font-size: 1.0em;" class="medium button red" width="150px" onclick="saverevisi();"><i class="fa fa-arrow-left" aria-hidden="true"></i>&nbsp;&nbsp;Revisi</a>
					</td>
					<td style="padding-left: 0px; text-align: left;" width="180px">
						<a id="updatebtn" href="#" class="medium button green" onclick="saveedit();"><i class="fa fa-save" aria-hidden="true"></i>&nbsp;&nbsp;Simpan</a>
						<a id="aanbtn" href="#" class="medium button green" onclick="saveedit();"><i class="fa fa-save" aria-hidden="true"></i>&nbsp;&nbsp;Simpan</a>
						<a id="approvebtn" href="#" class="medium button green" onclick="approve();"><i class="fa fa-thumbs-up" aria-hidden="true"></i>&nbsp;&nbsp;Approve</a>
					</td>
					<td style="padding-left: 0px; text-align: left;" width="400px">
						<a id="revisibtn" href="#" class="medium button red" width="150px" onclick="saverevisi();"><i class="fa fa-arrow-left" aria-hidden="true"></i>&nbsp;&nbsp;Revisi</a>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr height="0px">
		<td colspan="4">
			<input type="hidden" id="emohonid" name="emohonid">
			<input type="hidden" id="enostep" name="enostep">
			<input type="hidden" id="eposisi" name="eposisi">
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

<form method="post" name="filterfrm" id="filterfrm" action="permohonan.aspx" enctype="multipart/form-data">
<table id="filtertab" style="text-align: left; margin-left: 270px; margin-right: auto; margin-top: auto;" width="1160px" border="0" cellspacing="0" cellpadding="0" bgcolor="#E7FCFB">
	<tr height="20px">
		<td colspan="4" style="padding-left: 20px; text-align: left; font-size: 1.0em">
		</td>
	</tr>
	<tr height="38px">
		<td style="padding-left: 0px; padding-left: 30px; text-align: left; font-size: 0.9em; vertical-align: middle" width="190px">
			<FONT face="Arial" color="#000">
				Periode 
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 0.9em; vertical-align: middle" width="40px">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: middle" width="930px">
			<FONT face="Arial" color="#000">
				<!--input style="font-family: Arial; font-size: 0.9em" type="text" id="datepicker-ex30" name="editbbmtglbayar" class="textboxtgl"-->
				<input autocomplete="off" STYLE="padding-left:5px; font-size:0.9em; background-color: #FFFFFF; border:1px solid #ababab; height: 24px; width: 120px" autocomplete="off" class="tcal" id="tglperiode1" name="tglperiode1" type="text" value="<% response.write(tgl1) %>" tabindex="1">
			&nbsp;&nbsp; s/d &nbsp;&nbsp;
				<input autocomplete="off" STYLE="padding-left:5px; font-size:0.9em; background-color: #FFFFFF; border:1px solid #ababab; height: 24px; width: 120px" autocomplete="off" class="tcal" id="tglperiode2" name="tglperiode2" type="text" value="<% response.write(tgl2) %>" tabindex="1">
			</FONT>
		</td>
	</tr>
	<tr height="38px">
		<td style="padding-left: 30px; text-align: left; font-size: 0.9em; vertical-align: middle">
			<FONT face="Arial" color="#000">
				Nama Pengadaan
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 0.9em; vertical-align: middle">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; font-family: Arial; font-size: 1.0em; vertical-align: middle">
			<FONT face="Arial" color="#000">
				<!--select id="barutpmobil" name="barutpmobil" STYLE="color: #000000; background-color: #ffffff; font-size: 1.2em; width:320px"-->
				<div>
					<!--input style="padding-right: 10px; text-align: left; font-family: Arial; font-size: 1.2em" type="text" id="brgmerk" name="brgmerk" class="textbox"-->
					<input style="padding-right: 10px; text-align: left; width: 500px; height: 20px" type="text" id="namamohon" name="namamohon" class="textbox" value="<% Response.Write(namamohon) %>">
				</div>
			</FONT>
		</td>
	</tr>
	<tr height="38px">
		<td style="padding-left: 30px; text-align: left; font-size: 0.9em; vertical-align: middle">
			<FONT face="Arial" color="#000">
				Tahapan Proses
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.0em; vertical-align: middle">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: middle">
			<FONT face="Arial" color="#000">
			<select STYLE="color: #000000; background-color: #ffffff; font-size: 0.9em; width: 220px" id="soflow1" name="tahapan" data-placeholder="Jenis" tabindex="2"> 
				<!--option value="" >&nbsp;&nbsp;</option-->
				<option value="0" 
<% 	If Val(tahap)=0 Then
		response.write("selected")
	End If
%>
					>Semua</option>
				<option value="1"
<% 	If Val(tahap)=1 Then
		response.write("selected")
	End If
%>				
					>Permohonan</option>
				<option value="2"
<% 	If Val(tahap)=2 Then
		response.write("selected")
	End If
%>				
					>Evaluasi Kelengkapan</option>
				<option value="3"
<% 	If Val(tahap)=3 Then
		response.write("selected")
	End If
%>				
					>Proses ACC</option>
				<option value="4"
<% 	If Val(tahap)=4 Then
		response.write("selected")
	End If
%>				
					>Panggil Vendor</option>
				<option value="5"
<% 	If Val(tahap)=5 Then
		response.write("selected")
	End If
%>				
					>Aanwijzing</option>
				<option value="6"
<% 	If Val(tahap)=6 Then
		response.write("selected")
	End If
%>				
					>Nego Harga</option>
				<option value="7"
<% 	If Val(tahap)=7 Then
		response.write("selected")
	End If
%>				
					>Checklist Uji Kepatuhan</option>
				<option value="8"
<% 	If Val(tahap)=8 Then
		response.write("selected")
	End If
%>				
					>Surat Penunjukan</option>
				<option value="9"
<% 	If Val(tahap)=9 Then
		response.write("selected")
	End If
%>				
					>SPK</option>
				<option value="10"
<% 	If Val(tahap)=10 Then
		response.write("selected")
	End If
%>				
					>PKS</option>
			</select>			
			</FONT>
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
			<a id="clearbtn" href="#" class="medium button red" onclick="reset();">Reset</a>
		</td>
	</tr>
	<tr height="20px">
		<td colspan="3" style="padding-left: 20px; text-align: left;">
			<input type="hidden" id="userid" name="userid" value="">
			<input type="hidden" id="pageno" name="pageno" value="1">
			<input type="hidden" id="statusfilter" name="statusfilter" value="0">
		</td>
	</tr>
</table>

<table style="text-align: left; margin-left: 270px; margin-right: auto; margin-top: auto;" width="1160px" border="0" cellspacing="0" cellpadding="0">
	<tr height="30px">
		<td colspan="4" style="padding-left: 20px; text-align: left; font-size: 1.0em">
		</td>
	</tr>
</table>

<table id="mybar" style="margin-left: 270px;" width="1160px" align="center" cellpadding="2" cellspacing="2" border="0" bgcolor="#EEE">
	<tr height="38px" align="left">
		<td style="font-family: Arial; font-size: 0.9em;text-align: left;" width="120px">
			<!--a id="barubbm" href="" class="small button blue" onclick="addkct(); return false;"><i class="fas fa-plus fa-sm"></i>&nbsp;&nbsp;New</a-->
<%
			If Val(posisiq)<=15 Then
%>
			<a id="databarubtn" style="text-decoration: none" href="#" class="css-button-3" onclick="databaru();">
				<span class="css-button-3-icon"><i class="fa fa-plus" aria-hidden="true" id="i_pos"></i></span>
				<span class="css-button-3-text" onclick="databaru();"><FONT face="Tahoma" size="2">Data Baru</FONT></span>
			</a>
<%
			End If
%>
		</td>
		<td style="font-family: Arial; font-size: 0.9em;text-align: left;" width="900px">
		</td>
		<td style="font-family: Arial; font-size: 0.9em;text-align: left;" width="140px">

		</td>
	</tr>
</table>
<table id="mylist" style="margin-left: 270px;" width="1160px" align="center" cellpadding="2" cellspacing="2" border="0" bgcolor="#FAFAFA" >
	<tr height="36px" align="left" style="background-color:#D5D5D5;color:Black;" >
		<td style="font-family: Arial; font-size: 0.9em;text-align: center;" width="50px"><b>NO</b></td>                        
		<td style="font-family: Arial; font-size: 0.9em;text-align: center;" width="110px"><b>TGL</b></td>            
		<td style="font-family: Arial; font-size: 0.9em;text-align: center;" width="360px"><b>NAMA PENGADAAN</b></td>
		<td style="font-family: Arial; font-size: 0.9em;text-align: center;" width="140px"><b>JML ANGGARAN</b></td>
		<td style="font-family: Arial; font-size: 0.9em;text-align: center;" width="50px"><b>NI</b></td>
		<td style="font-family: Arial; font-size: 0.9em;text-align: center;" width="50px"><b>IP</b></td>
		<td style="font-family: Arial; font-size: 0.9em;text-align: center;" width="50px"><b>TOR</b></td>
		<td style="font-family: Arial; font-size: 0.9em;text-align: center;" width="50px"><b>EV</b></td>
		<td style="font-family: Arial; font-size: 0.9em;text-align: center;" width="140px"><b>TAHAP</b></td>
<%
		If posisint < 4 OR posisint > 13 Then
%>
		<td style="font-family: Arial; font-size: 0.9em;text-align: center;" width="90px"><b>VIEW/EDIT</b></td>
		<td style="font-family: Arial; font-size: 0.9em;text-align: center;" width="70px"><b>HAPUS</b></td>
<%
		Elseif posisint > 3 Then
%>
		<td style="font-family: Arial; font-size: 0.9em;text-align: center;" width="180px"><b>VIEW</b></td>
<%
		End If
%>
				
	</tr>

<% 

	Dim folderx as String
	Dim filesuratx as String
	Dim fileipx as String
	Dim filetorx as String
	Dim filerabx as String
	Dim i as Integer
	Dim tglsu as String
	Dim tglsua as String = ""
	Dim tglsub as String = ""
	Dim tglsuc as String = ""
	Dim tglxu as String
	Dim tglar(10) as String
	Dim spi as Integer = 0
	Dim ipi as Integer = 0
	Dim tori as Integer = 0
	Dim rabi as Integer = 0
	Dim spx as String = ""
	Dim ipx as String = ""
	Dim torx as String = ""
	Dim rabx as String = ""
	Dim posisir as Integer = 0

'==========================================================================
	posisir = Val(posisiq)

	If posisir=6 or posisir=7 or posisir=11 or posisir=12 or posisir=13 Then
		klpstr = ""
	End If

	'SQL2 = "SELECT ID FROM pengadaan WHERE (ID>0" & namastr & tglstr & tahapstr & klpstr & divstr & ") ORDER BY tglmohon DESC"
	'SQL1 = "SELECT ID FROM pengadaan WHERE (ID>0" & namastr & tglstr & tahapstr & ") ORDER BY tglmohon DESC"
	SQL2 = "SELECT * FROM [pengadaan] WHERE (ID>0) ORDER BY tglmohon DESC"

	myCmd.CommandText = SQL2
	myConn.Open()
	myReader = myCmd.ExecuteReader()
	
	a = 0
	jmldata = 0
	If myReader.HasRows Then
		While myReader.Read()
			a = a + 1
		End While
	End If
	myConn.Close()
	jmldata = a
			

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

	startval = ((hal-1)*20) + 1
	endval = startval + rowsperpage - 1
'*********

	aaa = 1
	ii = pagerows
	k = 1
	a = endval - startval
	bb = 0
	
    Try
        If aaa=1 Then
'============================================================================
' Data PENGADAAN
'============================================================================
			'SQL2 = "SELECT ID FROM [pengadaan] WHERE (ID>0" & namastr & tglstr & ") ORDER BY tglmohon DESC"
			SQL2 = "SELECT ID FROM pengadaan WHERE (ID>0" & namastr & tglstr & tahapstr & klpstr & divstr & ") ORDER BY tglmohon DESC"
			SQL2 = "SELECT ID FROM [pengadaan] WHERE (ID>0" & namastr & ") ORDER BY tglmohon DESC"

			myCmd.CommandText = SQL2
			myConn.Open()
			myReader = myCmd.ExecuteReader()
			
			a = 0
			b = 0
			jmldata = 0
			If myReader.HasRows Then
				While myReader.Read()
					a = a + 1
					If a>=startval AND  a<=endval Then
						b = b + 1
						idz(b) = myReader("ID").ToString()
					End If
					'tglzz(a) = myReader("tglsurat1").ToString()
				End While
			End If
			myConn.Close()
			jmldata = a

'response.write("---------------------------------------------------" & startval & "+" & endval & SQL2)	
'==========================================================================
	
			c = 0
			for i = startval to endval
				ii = ii + k
				bb = bb + 1
				c = c + 1
				SQL3 = "SELECT * FROM pengadaan WHERE ID=" & idz(c)

				myCmd.CommandText = SQL3
				myConn.Open()
				myReader = myCmd.ExecuteReader()
				
				a = 0
				jmldata = 0
				If myReader.HasRows Then
					While myReader.Read()
						a = a+1
						idx = myReader("id").ToString()
						nosuratx = myReader("kode").ToString()
						tglx = myReader("tglmohon").ToString()
						tahunx = myReader("tahun").ToString()
						jenisx = myReader("jenispengadaan").ToString()
						jenisangx = myReader("jenisanggaran").ToString()
						namax = myReader("nama").ToString()
						nilaix = Val(myReader("nilai").ToString())
						cuserx = myReader("createduser").ToString()
						nostepx = myReader("nostep").ToString()
						nostepijinx = myReader("nostepijin").ToString()
						stepx = myReader("step").ToString()
						statusrevisix = myReader("statusrevisi").ToString()
						folderx = myReader("namafolder").ToString()
						filesuratx = myReader("namafilesurat").ToString()
						fileipx = myReader("namafileip").ToString()
						filetorx = myReader("namafiletor").ToString()
						filerabx = myReader("namafilerab").ToString()
						If filesuratx.length() > 3 Then
							spi = 1
						End If
						If fileipx.length() > 3 Then
							ipi = 1
						End If
						If filetorx.length() > 3 Then
							tori = 1
						End If
						If filerabx.length() > 3 Then
							rabi = 1
						End If
						
						stepstr = stepx
						If stepx="" or stepx.ToUpper()="IJIN" Then
							If Val(nostepijinx)=1 Then
								stepstr = "Permohonan/Usulan (1)"
							ElseIf Val(nostepijinx)=2 Then
								stepstr = "Appr Pengelola (2)"
							ElseIf Val(nostepijinx)=3 Then
								stepstr = "Appr Pemp Kel (3)"
							ElseIf Val(nostepijinx)=4 Then
								stepstr = "Appr DGM (4)"
							ElseIf Val(nostepijinx)=5 Then
								stepstr = "Appr GM (5)"
							End If
						Else
							stepstr = "PROSES"
						End If
					End While
				End If
				myConn.Close()
				jmldata = a	
				
				If Val(statusrevisix)=1 Then
					stepstr = "Revisi"
				End If
				
				SQL3 = "SELECT * FROM pengadaan_revisi WHERE pengadaanID=" & idz(c)

				myCmd.CommandText = SQL3
				myConn.Open()
				myReader = myCmd.ExecuteReader()

				If myReader.HasRows Then
					While myReader.Read()
						alasanx = myReader("alasan").ToString()
					End While
				End if
				myConn.Close()
				
				Dim dt As DateTime = Convert.ToDateTime(tglx)
				tglsu = format(dt, "d MMM yyyy")
				tglqu = format(dt, "d-M-yyyy")
				tglsua = format(dt, "dd").ToString()
				tglsub = format(dt, "MMM").ToString()
				tglsuc = format(dt, "yyyy").ToString()
				If tglsub = "May" Then
					tglsub = "Mei"
				Elseif tglsub = "Aug" Then
					tglsub = "Agu"
				Elseif tglsub = "Oct" Then
					tglsub = "Okt"
				Elseif tglsub = "Dec" Then
					tglsub = "Des"
				End If
				tglsu = tglsua & " " & tglsub & " " & tglsuc
				tglxu = format(dt, "yyyy-M-d")
				'Dim dt As DateTime = Convert.ToDateTime(tglefektifx(i))
				'Dim dt As DateTime = format(tglefektifx(i), "yyyy-MM-dd")
				'tglefektifx(i) = dt
				'tglar = tglx.Split("-")
				'tglx = tglar(2) & "-" & tglar(1) & "-" & tglar(0)
				datastrx = idx & "," & tahunx & ",'" & jenisx & "','" & namax & "','" & tglxu & "','" & tglqu & "','" & jenisangx & "'," & nilaix & "," & _
								nostepx & "," & statusrevisix & ",'" & folderx & "','" & filesuratx & "','" & fileipx & _
								"','" & filetorx & "','" & filerabx & "'," & posisiq & ",'" & alasanx & "'"
				'datastrx(i) = idx(i) & "," & namax(i)
%>

	<tr id="row<% Response.Write(bb)%>" height="36px">
		<td style="padding-right: 5px; font-family: Arial; font-size: 0.9em;text-align: right; vertical-align: top;">
			<% Response.Write (i) %>
		</td>
		<td style="padding-left: 5px; padding-right: 20px; font-family: Arial; font-size: 0.9em;text-align: right; vertical-align: top;">
			<% Response.Write (tglqu) %>
		</td>
		<td style="padding-left: 5px; padding-bottom: 4px; font-family: Arial; font-size: 0.9em;text-align: left; vertical-align: top; line-height: 120%;">
			<% Response.Write (namax.ToUpper()) %>
		</td>
		<td style="padding-right: 5px; font-family: Arial; font-size: 0.9em;text-align: right; vertical-align: top;">
			<% Response.Write (FormatNumber(nilaix,0)) %>
		</td>
		<td style="padding-left: 5px; font-family: Arial; font-size: 0.9em;text-align: center; vertical-align: top;">
<% 
	If spi=0 Then
%>
		<img id="spi<% response.write(i) %>" src="./images/exit.png" height="20px" width="20px"/>
<%
	Elseif spi=1 Then
%>
		<img id="spi<% response.write(i) %>" src="./images/checkme.png" height="20px" width="20px"/>
<%
	End If
%>
		</td>
		<td style="padding-left: 5px; font-family: Arial; font-size: 0.9em;text-align: center; vertical-align: top;">
<% 
	If ipi=0 Then
%>
		<img id="ipi<% response.write(i) %>" src="./images/exit.png" height="20px" width="20px"/>
<%
	Elseif ipi=1 Then
%>
		<img id="ipi<% response.write(i) %>" src="./images/checkme.png" height="20px" width="20px"/>
<%
	End If
%>
		</td>
		<td style="padding-left: 5px; font-family: Arial; font-size: 0.9em;text-align: center; vertical-align: top;">
<% 
	If tori=0 Then
%>
		<img id="tori<% response.write(i) %>" src="./images/exit.png" height="20px" width="20px"/>
<%
	Elseif tori=1 Then
%>
		<img id="tori<% response.write(i) %>" src="./images/checkme.png" height="20px" width="20px"/>
<%
	End If
%>
		</td>
		<td style="padding-left: 5px; font-family: Arial; font-size: 0.9em;text-align: center; vertical-align: top;">
<% 
	If rabi=0 Then
%>
		<img id="rabi<% response.write(i) %>" src="./images/exit.png" height="20px" width="20px"/>
<%
	Elseif rabi=1 Then
%>
		<img id="rabi<% response.write(i) %>" src="./images/checkme.png" height="20px" width="20px"/>
<%
	End If
%>
		</td>
<%
	If Val(statusrevisix)=1 Then
%>
		<td style="padding-left: 0px; font-family: Arial; font-size: 0.9em;text-align: center; color: #FFF; background-color: #EF5858; vertical-align: top;">
<%
	ElseIf Val(statusrevisix)=0 Then
%>
		<td style="padding-left: 0px; font-family: Arial; font-size: 0.9em;text-align: center; color: #000; background-color: #FFF; vertical-align: top;">
<%
	End If
%>
<% 
	If stepx.ToUpper()="VER" Then
		stepstr = "VERIF (" & nostepx & ")"
	End If
	Response.Write (stepstr) 
%>
		</td>
		<td style="padding-left: 0px; font-family: Arial; font-size: 0.9em;text-align: center; vertical-align: top;">
			<!--a id="edit<% Response.Write (i)%>" href="" class="small button blue" onclick="addkct(); return false;"><i class="fas fa-edit fa-sm"></i></a-->
<%
	If stepstr="IJIN" Then
%>
			<a id="edit<% Response.Write (i)%>" href="#" onclick="editdataxx(<% Response.Write(datastrx) %>);">
			<img id="uploadPreview<% Response.Write (i)%>" src="./images/edit2.jpg" height="26px" width="25px"/><br/>
			</a>
<%
	End If
%>
		</td>

		<td style="padding-left: 0px; font-family: Arial; font-size: 0.9em;text-align: center; vertical-align: top;">
<%
		If (posisiq < 4 OR posisiq > 13) AND nostepx=1 AND statusrevisix=0 Then
%>
			<!--a id="hapus<% Response.Write(i)%>" href="#" class="small button red" onclick="addkct(); return false;"><i class="fas fa-trash fa-sm"></i></a-->		
			<img id="uploadPreview<% Response.Write (i)%>" src="./images/delete4.png" OnClick="hapusada(<% Response.Write (idx) %>);" height="26px" width="25px"/><br/>
<%
		End If
%>
		</td>

	</tr>
<%
			Next i
        End If

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

<%
'response.write("-----------------------------------------" + klpstr + divstr + "<br>")
'response.write("-----------------------------------------" + SQL1)
%>

<table id="paging" style="text-align: left; margin-left: 270px; margin-right: auto; margin-top: auto;" width="1100px" border="0" cellspacing="0" cellpadding="0">
	<tr height="34px" style="background-color: #FFFFFF">
		<td style="padding-left: auto; text-align: left; font-size: 0.9em;" width="120px">
			<div class="pagination">
			  <a href="#" onclick="viewpage(<% Response.Write(pageno-1) %>)">&laquo;</a>
<%
			for a=1 to jmlpage
%>
			  <a 
			  <% If a=pageno Then 
					Response.Write ("class='active'")
				 End If
			  %>
				href="#" onclick="viewpage(<% Response.Write(a) %>)"><% Response.Write(a) %></a>
<%
			next a
%>
			  <a href="#" onclick="viewpage(<% Response.Write(pageno+1) %>)">&raquo;</a>
			</div>	
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
			Session("namaklp") = namaklpq
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


<form name="pagefrm" id="pagefrm" method="post" action="main.aspx" enctype="multipart/form-data">
	<input type="hidden" id="muserid" name="muserid" value="<%  %>">
	<input type="hidden" id="pageno" name="pageno">
</form>

<form name="approvefrm" id="approvefrm" method="post" action="pengadaanapprovego.aspx" enctype="multipart/form-data">
	<input type="hidden" id="anostep" name="anostep">
	<input type="hidden" id="aposisi" name="aposisi">
	<input type="hidden" id="aadaid" name="aadaid">
</form>

<form name="myeditfrm" id="myeditfrm" method="post" action="permohonanedit.aspx" target="_blank" enctype="multipart/form-data">
	<input type="hidden" id="edituserid" name="edituserid">
	<input type="hidden" id="padaid" name="padaid">
	<input type="hidden" id="statusmohon" name="statusmohon" value="0">
</form>

<form name="revisifrm" id="revisifrm" method="post" action="pengadaanrevisigo.aspx" enctype="multipart/form-data">
	<input type="hidden" id="rnostep" name="rnostep">
	<input type="hidden" id="ralasan" name="ralasan">
	<input type="hidden" id="rposisi" name="rposisi">
	<input type="hidden" id="radaid" name="radaid">
</form>

<form method="post" name="delfrm" id="delfrm" action="pengadaandelgo.aspx" enctype="multipart/form-data">
	<input type="hidden" id="deladaid" name="deladaid">
</form>


</body>
</html>
