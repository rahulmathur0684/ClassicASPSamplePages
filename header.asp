<%
If(Request.QueryString("untitle")<>"") then
Response.Status="301 Moved Permanently"
Response.AddHeader "Location",pathOfdata&"/"
Response.End()
End if
If LCase(Request.ServerVariables("UNENCODED_URL"))="/home" then
Response.Status="301 Moved Permanently"
Response.AddHeader "Location",pathOfdata&"/"
Response.End()
End If
If LCase(Request.ServerVariables("UNENCODED_URL"))="/home/" then
Response.Status="301 Moved Permanently"
Response.AddHeader "Location",pathOfdata&"/"
Response.End()
End If
%>
<!DOCTYPE html>
<html lang="en">
<head>
    <% 
    Response.charset="utf-8"
    Response.CacheControl = "no-cache, no-store, must-revalidate"
    Response.AddHeader "X-FRAME-OPTIONS", "DENY" 
    %>
    <!--#include file="MobileSiteRedirectionNew.asp" -->
    <!--#include file="Connections/config.asp" -->
    <!--#include file="RelCanonicalurl.asp" -->
    <!--#include file="HtmlDecode.asp" -->
    <!--#include file="AddCartHeader.asp" -->
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1, minimum-scale=1, user-scalable=0">
    <meta name="p:domain_verify" content="19406e4bea75016269640da444122d84">
    <!--#include file="MetaDesc.asp" -->
    <!-- meta start from Here -->
    <title><% If RStour__MMColParam = 0 Then %><%=titletag %><% else %><%=titlefromdatabase%><% end if %></title>
    <meta name="description" content="<% If RStour__MMColParam = 0 Then %><%=metatag %><% else %><%=SEOdesc%><% end if %>" />
    <% if Isnoindexnofollow= "N" then %>
    <meta name="robots" content="noindex, nofollow">
    <%Else%>
    <meta name="robots" content="NOYDIR, NOODP" />
    <meta name="robots" content="index, follow" />
    <% End if %>
    <meta property="og:locale" content="en_US" />
    <meta property="og:type" content="website" />
    <meta property="og:title" content="<% If RStour__MMColParam = 0 Then %><%=titletag %><% else %><%=titlefromdatabase%><% end if %>">
    <meta property="og:description" content="<% If RStour__MMColParam = 0 Then %><%=metatag %><% else %><%=SEOdesc%><% end if %>" />
    <meta property="og:url" content="https://www.xyz.com<%=CurrentUrl%>" />
    <meta property="og:site_name" content="xyz" />
    <meta property="Amalfi Coast:location:latitude" content="40.6333329" />
    <meta property="Amalfi Coast:location:longitude" content="14.601766" />
    <meta property="Italy:locality" content="Amalfi Coast" />
    <meta name="twitter:card" content="summary" />
    <meta name="twitter:url" content="https://www.xyz.com<%=CurrentUrl%>" />
    <meta name="twitter:title" content="<% If RStour__MMColParam = 0 Then %><%=titletag %><% else %><%=titlefromdatabase%><% end if %>" />
    <meta name="twitter:description" content="<% If RStour__MMColParam = 0 Then %><%=metatag %><% else %><%=SEOdesc%><% end if %>" />
    <% if CurrentUrl= "/festivals-in-italy" then %>
    <meta name="twitter:image" content="https://www.xyz.com/images/festivalimages/Capodanno,%20Italy%20-1.jpg">
    <% End if %>
    <meta name="twitter:site" content="@xyzTours" />
    <meta name="twitter:domain" content="xyz Limos and Tours" />
    <meta name="twitter:creator" content="@xyzTours" />
    <meta property="article:publisher" content="" />
    <meta property="Amalfi Coast:location:latitude" content="40.6333329" />
    <meta property="Amalfi Coast:location:longitude" content="14.601766" />
    <meta property="Italy:locality" content="Amalfi Coast" />
    <meta property="og:site_name" content="xyz" />
    <meta name="FbImage" property="og:image" id="metaFBImage" content="https://www.xyz.com/images/tours/Backgroundimages/image1.jpg" />
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
    <meta name="author" content="xyz Limos" />
    <meta name="dcterms.rightsHolder" content="xyz Limos" />
    <link rel="publisher" href="" />
    <meta name="DC.title" content="Italy Tours | Private Italian Tours | Private Guided Day Tours Italy" />
    <meta name="geo.region" content="IT-RM" />
    <meta name="geo.placename" content="Praiano" />
    <meta name="geo.position" content="40.609966;14.532112" />
    <meta name="ICBM" content="40.609966, 14.532112" />
    <meta name="msvalidate.01" content="CB3FFE0543B4BA35D1BB2D50630E2B05" />
    <meta name="google-site-verification" content="6aprO9JyRdznpLOVwJHCsQgyMKPWP8kCWDdDff8qUNY" />
    <!-- meta end from Here -->

    <!-- Bootstrap -->
    <link href="<%=pathOfdata %>/css/bootstrap.min.css" rel="stylesheet">
    <link href="<%=pathOfdata %>/css/owl.carousel.min.css" rel="stylesheet">
    <link href="<%=pathOfdata %>/css/customscrollbar.css" rel="stylesheet">
    <link href="<%=pathOfdata %>/css/animate.min.css" rel="stylesheet">
    <link href="<%=pathOfdata %>/css/hover.css" rel="stylesheet">
    <link href="<%=pathOfdata %>/boot/alertify.core.css" rel="stylesheet">
    <link href="<%=pathOfdata %>/boot/alertify.default.css" rel="stylesheet">
    <link href="<%=pathOfdata %>/css/reset.css" rel="stylesheet">
    <link href="<%=pathOfdata %>/css/style.css" rel="stylesheet">
    <link href="<%=pathOfdata %>/css/responsive.css" rel="stylesheet">
    <link href="<%=pathOfdata %>/css/font-awesome.min.css" rel="stylesheet">
    <link href="<%=pathOfdata %>/css/material-design-iconic-font.min.css" rel="stylesheet">
    <link rel="stylesheet" type="text/css" href="<%=pathOfdata %>/jquery.datepick.package-4.1.0/jquery.datepick.css">
    <link rel="stylesheet" href="https://code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
    <link href="<%=pathOfdata %>/fancyBox-master/source/jquery.fancybox.css" rel="stylesheet" />
    <script src="<%=pathOfdata %>/js/jquery.min.js"></script>
    <script src="<%=pathOfdata %>/js/AllJS.js"></script>
    <script src="<%=pathOfdata %>/js/jssor.slider.min.js"></script>
    <script src="https://www.google.com/recaptcha/api.js" async defer></script>
    <script id="Cookiebot" src="https://consent.cookiebot.com/uc.js" data-cbid="ce0ca6d5-4bd4-41cc-9ec4-c706857f743d" type="text/javascript" async></script>
    <script id="CookieDeclaration" src="https://consent.cookiebot.com/ce0ca6d5-4bd4-41cc-9ec4-c706857f743d/cd.js" type="text/javascript" async></script>


</head>
<body>
    <!-- Google Tag Manager (noscript) -->
    <noscript>
        <iframe src="https://www.googletagmanager.com/ns.html?id=GTM-MWHBWLC"
            height="0" width="0" style="display: none; visibility: hidden"></iframe>
    </noscript>
    <!-- End Google Tag Manager (noscript) -->
    <main>
        <header>
            <nav class="navbar navbar-default navbar-fixed-top">
                <div class="container-fluid">
                    <div class="navbar-header">
                        <button type="button" class="navbar-toggle collapsed" data-toggle="collapse" data-target="#navbar" aria-expanded="false" aria-controls="navbar"><span class="sr-only">Toggle navigation</span> <span class="navicon-bar"></span></button>
                        <a class="navbar-brand" href="<%=pathOfdata %>">
                            <img src="<%=pathOfdata %>/images/logo.png" alt="HeaderLogo" title="HeaderLogo"></a>
                    </div>
                    <div class="my-account-cart">
                        <% If IsNull(Session("LoginUser")) OR IsEmpty(Session("LoginUser")) OR Session("LoginUser")="" then  %>
                        <% if iToursLength>0 Then %>
                        <div class="cart-link">
                            <a href="<%=pathOfdata %>/shopping-cart.asp?remove=tourdata">
                                <figure>
                                    <img src="<%=pathOfdata %>/images/cart-icn.png" alt="Cart" title="Cart">
                                    <figcaption><%=iToursLength%></figcaption>
                                </figure>
                            </a>
                        </div>
                        <% End IF %>
                        <div class="login-reg-link">
                            <a href="<%=pathOfdata %>/login/">
                                <svg enable-background="new 0 0 512 512" id="my-account" version="1.1" viewBox="0 0 512 512" xml:space="preserve" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink">
                                    <g>
                                        <path d="M256,274.1c-75.6,0-137.1-61.5-137.1-137.1S180.4,0,256,0c75.6,0,137.1,61.5,137.1,137.1   S331.6,274.1,256,274.1z M256,20.4c-64.3,0-116.6,52.3-116.6,116.6S191.7,253.7,256,253.7c64.3,0,116.6-52.3,116.6-116.6   S320.3,20.4,256,20.4z" fill="#174a5a" />
                                        <path d="M493.9,512c-5.6,0-10.2-4.6-10.2-10.2c0-125.5-102.1-227.6-227.6-227.6c-125.5,0-227.6,102.1-227.6,227.6   c0,5.6-4.6,10.2-10.2,10.2s-10.2-4.6-10.2-10.2c0-66.3,25.8-128.6,72.7-175.4s109.2-72.7,175.4-72.7s128.6,25.8,175.4,72.7   s72.7,109.2,72.7,175.4C504.1,507.4,499.5,512,493.9,512z" fill="#174a5a" />
                                    </g>
                                </svg>
                                <span>Sign-in</span>
                            </a>
                        </div>
                        <% 
                        Else 
                        str = Session("LoginUserfirstname")          
                        FirstLetter = Left (str, 1)              
                        if(LCase(Request.ServerVariables("UNENCODED_URL"))<>"/edit-info/") then
                        %>
                        <!--#include file="forceuserupdate.asp" -->
                        <%end if %>
                        <div class="cart-link">
                            <a href="<%=pathOfdata %>/shopping-cart.asp?remove=tourdata">
                                <figure>
                                    <img src="<%=pathOfdata %>/images/cart-icn.png" alt="" />
                                    <figcaption><%=iToursLength%></figcaption>
                                </figure>
                            </a>
                        </div>
                        <div class="my-account-dd">
                            <div class="dropdown">
                                <button onclick="myFunction()" class="dropbtn">
                                    <span class="my-ac-dd"><%=UCase(FirstLetter)%></span><img class="close-icn-mc" style="display: none;" src="<%=pathOfdata %>/images/close-icn.png" alt="" />
                                </button>
                                <div id="myDropdown" class="dropdown-content content-1">
                                    <ul>
                                        <li><a href="<%=pathOfdata %>/account-info/">Account Information</a></li>
                                        <li><a href="<%=pathOfdata %>/mytestimonial/">My Testimonial</a></li>
                                        <li><a href="<%=pathOfdata %>/newmyservice.asp">My Services</a></li>
                                        <li>
                                            <button class="cmn-btn" onclick="userlogout();">Logout</button></li>
                                        <input type="hidden" id="hdnuserFName" name="hdnuserFName" value="<%=Session("LoginUserfirstname")%>" />
                                        <input type="hidden" id="hdnuserLName" name="hdnuserLName" value="<%=Session("LoginUserfirstname")%>" />
                                    </ul>
                                </div>
                            </div>
                        </div>
                        <% End IF %>
                    </div>
                    <div id="navbar" class="navbar-collapse collapse">
                        <h2 class="infonav hidden-xs hidden-sm">
                            <span><a href="tel:+390898424226 ">+39 089 842 4226 </a></span>
                            <span><a href="mailto:info@xyz.com">info@xyz.com</a> </span>
                        </h2>
                        <ul class="nav navbar-nav navbar-right hidden-xs hidden-sm">
                            <li class="classtest" id="ClasstestHeader">
                                <a href="javascript:void(0)" class="showmenu">Transfers <i class="fa fa-angle-down"></i></a>
                                <ul class="nav-sub-menu" style="display: none;">
                                    <li><a href="<%=pathOfdata %>/tour-enroute/" class="tournewclass">Tour Enroute</a></li>
                                    <li><a href="<%=pathOfdata %>/transfer/" class="tournewclass">Transfers A to B</a></li>
                                </ul>
                            </li>
                            <li><a href="<%=pathOfdata %>/tours/">City Tours</a></li>
                            <li><a href="<%=pathOfdata %>/shore-excursions/">Shore Excursions</a></li>
                            <li><a href="<%=pathOfdata %>/fleets/">Fleet</a></li>
                            <li><a href="<%=pathOfdata %>/blog/">Blog</a></li>
                            <li><a href="<%=pathOfdata %>/history/">Company</a></li>
                            <li><a href="<%=pathOfdata %>/faq/">FAQ</a></li>
                            <li><a href="<%=pathOfdata %>/contact/">Contact</a></li>
                            <li><a href="<%=pathOfdata %>/quote-request/">Quote Request</a></li>

                        </ul>
                        <div class="mobileNav show-m">
                            <div class="main-menu-top">
                                <ul class="nav navbar-nav">
                                    <li class="classtest"><a href="javascript:void(0)" class="showmenu">Transfers</a>
                                        <ul class="nav-sub-menu" style="display: none;">
                                            <li><a href="<%=pathOfdata %>/tour-enroute/">Tour Enroute</a></li>
                                            <li><a href="<%=pathOfdata %>/transfer/">Transfers A to B</a></li>
                                        </ul>
                                    </li>
                                    <li><a href="<%=pathOfdata %>/tours/">City Tours</a></li>
                                    <li><a href="<%=pathOfdata %>/shore-excursions/">Shore Excursions</a></li>
                                    <li><a href="<%=pathOfdata %>/fleets/">Fleet</a></li>
                                    <li><a href="<%=pathOfdata %>/quote-request/">Quote Request</a></li>
                                </ul>
                            </div>
                            <div class="nav-f-menu">
                                <ul class="others-m-list nav navbar-nav">
                                    <li><a href="#">Blog</a></li>
                                    <li><a href="<%=pathOfdata %>/history/">Company</a></li>
                                    <li><a href="<%=pathOfdata %>/faq/">FAQ</a></li>
                                    <li><a href="<%=pathOfdata %>/contact/">Contacts</a></li>
                                </ul>
                            </div>
                            <div class="call-now-top">
                                <span><a href="tel:+390898424226">+39 089 842 4226</a></span>
                                <span><i class="fa fa-envelope"></i><a href="mailto:info@xyz.com">info@xyz.com</a> </span>
                            </div>
                        </div>
                    </div>
                </div>
            </nav>
        </header>
