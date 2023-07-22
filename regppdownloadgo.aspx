<!DOCTYPE html>

<%@ Page Language="VB" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data.OleDb" %> 
<%@ Import Namespace="System.Configuration" %>
<%@ Import Namespace="System.Web.Configuration" %>
<%@ Import Namespace="System.IO" %>
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

<script language="vbscript" type="text/vbsscript" runat="server">

Function GetFullName$()
    Dim FirstName As String, LastName As String

    FirstName = "Sir"
    LastName = "Monay"

    Return LastName & ", " & FirstName
End Function

</script>

<title>Download</title>
</head>

<body>
<%
    'Dim FullName$

    'FullName = GetFullName()
    'Response.Write(GetFullName$)
%>

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
	

	Dim idq As String = ""
	Dim namauserq As String = ""
	Dim statusq As String = ""
	Dim nppq as String = ""
	Dim jabatanq As String = ""
	Dim klpq As String = ""
	Dim namaklpq as String = ""
	Dim unitq As String = ""
	Dim posisiq As String = ""
	Dim suratq As String = ""
	Dim sekreq As String = ""
	Dim adaq as String
	Dim asetq as String
	Dim statusaktifq as String = ""
	Dim divq As String = ""
	Dim levelaksesq as String = ""
	Dim posisix as String = ""
	Dim jmlsurat as Integer = 0
	
	Dim idxx(100000) as String
	Dim nosuratxx(100000) as String
	Dim tglxx(100000) as String
	Dim jenisxx(100000) as String
	Dim jenisstrxx(100000) as String
	Dim rubrikxx(100000) as String
	Dim perihalxx(100000) as String
	Dim tujuanxx(100000) as String
	Dim tembusanxx(100000) as String
	Dim jenisundanganxx(100000) as String
	Dim cuserxx(100000) as String
	Dim namafilexx(100000) as String
	Dim namafolderxx(100000) as String
	Dim nourutxx(100000) as Long
	Dim norekdebetxx(100000) as String
	Dim norekkreditxx(100000) as String
	Dim nominalxx(100000) as Long
	Dim klpxx(100000) as Integer
	Dim namaklpxx(100000) as String
	
	'Dim pageno as Integer = 1
	Dim rowperpage as Integer = 20
	Dim pageno as String
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

	Dim SQL0 As String
	Dim SQL1 As String
	Dim SQL2 As String
	Dim SQLUP as String = ""
	Dim SQLex as String = ""
	Dim writenama as String = ""
	Dim writenpp as String = ""
	
	Dim startval as Integer = 0
	Dim endval as Integer = 0
	Dim datastrx(5000) as String
	Dim datastrxx as String
	Dim perihal as String = ""
	Dim perihalstr as String = ""
	Dim tgl1 as String = ""
	Dim tgl2 as String = ""
	Dim tgl10 as string = ""
	Dim tgl20 as String = ""
	Dim tglstr as String = ""
	Dim tglsu as String
	Dim statusf as Integer = 0
	Dim i as Long
	Dim l as Integer
	Dim m as integer
	Dim n as Integer
	Dim jmlrubrik as Integer = 0
	Dim statusbum as Integer = 0
	Dim klpstr as String = ""
	Dim namauserstr as String = ""
	Dim statusbaru as Integer = 0
	
	
	Dim setmenu as String = "surat"
	
	Dim tahuniki as Integer
	Dim tahunwingi as Integer
	Dim tahunsesuk as Integer

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
%>

<!--#include file="koneksi.aspx"-->

<%	
    'Using workBook As New XLWorkbook(FileUpload1.PostedFile.InputStream)

			If NOT Request.Form("regvsql") is Nothing Then
				SQLex = Request.Form("regvsql")
			End If
		
			user0 = Session("username")
			idq = Session("userid")
			nppq = Session("idemployee")
			If not Session("namauser") is nothing Then
				namauserq = Session("namauser")
			End If
			jabatanq = Session("jabatan")
			klpq = Session("klp")
			If not Session("namaklp") is nothing Then
				namaklpq = Session("namaklp")
			End If
			unitq = Session("unit")
			posisiq = Session("posisi")
			posisix = Session("posisistr")
			suratq = Session("menusurat")
			sekreq = Session("menusekre")
			adaq = Session("menuada")
			asetq = Session("menuaset")
			divq = Session("divisi")
			levelaksesq = Session("levelakses")

			

			myCmd = myConn.CreateCommand

        'Read the first Sheet from Excel file.
			Dim exten As String = ""
			Dim namafile0 as String = ""
			Dim namafileout as String = ""
			Dim folder1 as String = "C:/Inetpub/wwwroot/portalbum/DOC/register/Temp/"
			'folder1 = ""
			'Dim folder1 as String = "http://localhost/portalbumDOC/surat/Temp/"
			Dim folder2 as String = "./DOC/register/Temp/"

			If Not System.IO.Directory.Exists(folder1) Then
				System.IO.Directory.CreateDirectory(folder1)
			End If

			namafile0 = folder1 & "regppOUT.xls"
			namafileout = folder1 & "regppOUT.xls"
			
			If Len(Dir$(namafileout)) > 0 Then Kill(namafileout)
			
			Dim strExcelConnr as String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & namafile0 & "';Extended Properties='Excel 12.0;HDR=Yes'"
			Dim strExcelConnw as String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & namafileout & "';Extended Properties='Excel 12.0;HDR=Yes'"


			Dim connExcel as new OleDbConnection(strExcelConnr)
			Dim cmdExcel as new OleDbCommand()
			cmdExcel.Connection = connExcel

			Dim connExcelw as new OleDbConnection(strExcelConnw)
			Dim cmdExcelw as new OleDbCommand()
			cmdExcelw.Connection = connExcelw

'==============================================================================================

			connExcel.Open()
			Dim dtExcelSchema as New DataTable()
			dtExcelSchema = connExcel.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Columns, Nothing)
			connExcel.Close()

			Dim datax as String
			Dim tglx(100000) as String
			Dim jenisx as String
			Dim perihalx as String
			Dim rubrikx as String
			Dim tujuanx as String
			Dim tembusanx as String
			Dim dt As DateTime
			Dim sqlout as String
			Dim oda As New OleDbDataAdapter()
			Dim ds as new DataSet()
			'Dim SheetName As String = dtExcelSchema.Rows(0)("TABLE_NAME").ToString()

			connExcelw.Open()
			'SQL0 = "CREATE TABLE Sheet2 ([NO] INTEGER PRIMARY KEY, [NOMOR] Text(255), [NOURUT] INTEGER, [TGL] DATE, [JENIS] INTEGER," & _
			'					"[PERIHAL] Text(255),[RUBRIK] Text(255),[TUJUAN] Text(255)," & _
			'					"[TEMBUSAN] Text(255), [UNDANGAN] INTEGER, [KELOMPOK] Text(255))"
			'SQL0 = "CREATE TABLE Sheet1 ([NO] INTEGER, [NOMOR] VARCHAR(200), [NOURUT] INTEGER, [TGL] DATE, [JENIS] INTEGER," & _
			'					"[PERIHAL] VARCHAR(200),[TUJUAN] VARCHAR(200)," & _
			'					"[TEMBUSAN] VARCHAR(200), [UNDANGAN] INTEGER, [KELOMPOK] VARCHAR(200))"
			SQL0 = "CREATE TABLE Sheet1 ([NO] INTEGER, [NOMOR] TEXT(200), [NOURUT] INT, [TGL] TEXT(100), [JENIS] TEXT(200)," & _
								"[PERIHAL] TEXT(200),[NOREK DEBET] TEXT(200),[NOREK KREDIT] TEXT(200)," & _
								"[NOMINAL] TEXT(200))"
'response.write(SQLex & "<br>")								
			cmdExcelw.CommandText = SQL0
			cmdExcelw.ExecuteNonQuery()
			'cmdExcelw.ExecuteReader()
			connExcelw.Close()
'==============================================================================================
			myCmd.CommandText = SQLex
'response.write(SQLex & "<br>")
			myConn.Open()
			myReader = myCmd.ExecuteReader()
			
			a = 0
			If myReader.HasRows Then
				While myReader.Read()
					a = a + 1
					jenisstrxx(a) = "BIASA"
					idxx(a) = myReader("ID").ToString()
					nosuratxx(a) = myReader("nomor").ToString()
					nourutxx(a) = myReader("nourut").ToString()
					tglxx(a) = myReader("tgl").ToString()
					klpxx(a) = Val(myReader("klp").ToString())
					rubrikxx(a) = myReader("rubrik").ToString()
					nominalxx(a) = myReader("nominal").ToString()
					cuserxx(a) = myReader("createduser").ToString()
					namafilexx(a) = myReader("namafile").ToString()
					namafolderxx(a) = myReader("namafolder").ToString()
					dt = Convert.ToDateTime(tglxx(a))
					'tglsu = dt.Day & " " & dt.Month & "  " & dt.Year
					tglx(a) = format(dt, "dd-MM-yyyy").ToString()
					'Dim dt As DateTime = format(tglefektifx(i), "yyyy-MM-dd")
				End While
			End If
			myConn.Close()
			jmlsurat = a
'response.write(jmlsurat & "<br>")			
			
			For b = 1 to jmlsurat
				SQL0 = "SELECT * FROM [klp] WHERE klp_id=" & klpxx(b)
				myCmd.CommandText = SQL0
				myConn.Open()
				myReader = myCmd.ExecuteReader()
				If myReader.HasRows Then
					While myReader.Read()
						namaklpxx(b) = myReader("klp_nama").ToString()
					End While
				End If
				myConn.Close()
			Next b
	'==============================================================================================
'response.write(jmlsurat & "<br>")
			For a = 1 to jmlsurat
				connExcelw.Open()
				cmdExcelw.Connection = connExcelw
				sqlout = "INSERT INTO [Sheet1$] ([NO], [NOMOR], [NOURUT], [TGL], [NOMINAL]) " & _
						"VALUES(" & a & ",'" & nosuratxx(a) & "'," & nourutxx(a) & ",'" & tglx(a) & "'," & nominalxx(a) & ")"
'response.write(sqlout & "<br>")
				cmdExcelw.CommandText = sqlout
				cmdExcelw.ExecuteNonQuery()
				connExcelw.Close()
				'response.write("<br>")
			Next a



	Dim namafileout1 as String = ""

	namafileout1 = folder2 + "surat.xls"

%>


<div class="example-element"></div>

<!--#include file="menu.aspx"-->

<!--#include file="topbar.aspx"-->


<div class="container">
			<div class="welcome">Selamat Datang,</div>
			<div class="photo"><img src="./images/user5.png" alt="" style="width:40px;height:40px;"></div>
			<div class="userin"><i><b><% Response.Write(namauserstr) %></b></i></div>
			<div class="posisi" STYLE="font-size: 0.8em"><i><% Response.Write(posisix) %></i></div>
</div>

<%
	namafileout = "./DOC/register/Temp/regppOUT.xls"
%>

<table id="datatab" style="margin-left: 270px;" align="center" width="800px" cellpadding="2" cellspacing="2" border="0" bgcolor="#FFFFFF" >
	<tr height="100px" align="left" style="background-color:#FFFFFF;color:Black;" >
		<td style="font-family: Arial; font-size: 0.9em;text-align: center;" width="200px"></td>                        
		<td style="font-family: Arial; font-size: 0.9em;text-align: center;" width="300px"></td>            
		<td style="font-family: Arial; font-size: 0.9em;text-align: center;" width="300px"></td>
	</tr>
	<tr height="100px" align="left" style="background-color:#FFFFFF;color:Black;" >
		<td style="font-family: Arial; font-size: 0.9em;text-align: center;" width="200px"></td>                        
		<td style="font-family: Arial; font-size: 0.9em;text-align: center;" width="300px">
			<a href="<% response.write(namafileout) %>" download rel="noopener noreferrer" target="_blank">
			   Download File
			</a>		
		</td>            
		<td style="font-family: Arial; font-size: 0.9em;text-align: center;" width="300px"></td>
	</tr>
</table>	
		
		
		


<%
        End If

    Catch ex As Exception
        'lbl1.Text="Error while connecting Server."
		'connExcelw.Close()
		Response.Write (ex.Message)

    Finally
		'myReader.Close()
        'myConn.Close() 'Whether there is error or not. Close the connection.

    End Try
%>

</body>
</html>