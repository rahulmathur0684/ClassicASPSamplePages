<%@  language="VBSCRIPT" codepage="65001" %>

<%''' This file used to redirect to new site %>

<!--#include file="NewPagesmap.asp" -->

<%
'***********In this page for topyourcart we already used "AjaxYourCardtesting07032012"************
'*****and for page we are including "customize/EditReservationCustome.asp" to show custome details in page*****
%>

<!--#include file="Connections/AjaxYourCardtesting07032012.asp"-->

<%
'***In the AjaxYourCard we have include the config file***
Dim dbName
dbName = "dbo213093820."
' or if you actually want dd-monthname-YYYY instead of d-monthname-YYYY
Function PadLeft(Value, Digits)
   PadLeft = CStr(Value)
   If Len(PadLeft) < Digits Then
      PadLeft = Right(String(Digits, "0") & PadLeft, Digits)
   End If
End Function

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN">
<html>
<head>

    <title>Limos chauffeur service taxi excursions Amalfi Coast, amalfi coast tours, Positano,
        Amalfi, Ravello, Sorrento, amalfi coast excursions, Rome, Florence, excursions,
        pompeii, transfers, rome tours, rome excursions</title>

    <!--'******If Image is spreading in IE or etc happening,then use this line "Just after title tag"****************-->
    <meta http-equiv="X-UA-Compatible" content="IE=8;FF=3;Opera=9;Konqueror=3;Safari=3" />
    <!--'************End of this line**************-->

    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />


    <link href="Calender.css" rel="stylesheet" type="text/css" />

    <link rel="shortcut icon" href="favicon-1.ico" />

    <script type="text/javascript" src="js/utilities.js"></script>

    <script type="text/javascript" src="js/Common.js"></script>

    <script type="text/javascript" src="js/jquery.js"></script>

    <script type="text/javascript" src="js/Calender_duplicate.js"></script>

    <script language="JavaScript1.2" type="text/javascript" src="mm_css_menu.js"></script>

    <style type="text/css" media="screen">
        @import url("./mainpage.css");
    </style>
</head>
<%  
'***Get Cookies  values to Add tours in the cart
Dim sCookieValue,arrTours,iToursLength,sCookieTransferService,arrytransferTourList,itransferLength
'****************14-04-2012******************
Dim MessagetoSave 'If no booking for this user occures,then we save it add it on 14-04-2012
MessagetoSave="Success"
'**********************************
itransferLength=0
iToursLength=0
sCookieValue = Request.Cookies("toursbooked")
arrTours = Split(sCookieValue,"||")
iToursLength = UBound(arrTours) + 1

sCookieTransferService = Request.Cookies("transfertours")
arrytransferTourList=Split(sCookieTransferService,"||")
itransferLength = UBound(arrytransferTourList) + 1

%>
<body bgcolor="#eadf95" onload="MM_preloadImages('images/home_f2.jpg','images/history_f2.jpg','images/transfers_f2.jpg','images/shore_excursions_f2.jpg','images/tours_f2.jpg','images/reservations_f2.jpg','images/fleet_f2.jpg','images/contact_email_f2.jpg','images/testimonials_f2.jpg');">
    <div id="FWTableContainer2135516502">
        <div id="contentcontainer" style="background-image: url(images/header.jpg); height: 79px;">
            <!--Start code for the top your cart section This is the part of the Ajax-->
            <% call TopYourCart()%>
            <!--End Top your cart section DIV-->
        </div>

        <!--#include file="topnav3.asp" -->
        <div id="bgrnd_reservations">
            <div style="width: 700px; position: relative; left: 140px;">
                <br />
                <%
                
            '***********************************Add this on 04-03-2012*************************************************************
            'Here the code for show which color is used for which purpose           
            dim mysysIp
            mysysIp= Request.ServerVariables("REMOTE_ADDR")
			 
			 if(InStr(mysysIp,"183.182")="1") then         
                
                Response.Write("<br/>" & Request.ServerVariables("HTTP_REFERER"))'o/p:- http://www.xyz.com/reservations_testing02042012.asp                
                Response.Write("<br/>" & Request.ServerVariables("REQUEST_METHOD")) 'o/p:- GET or POST
                Response.Write("<br/>" & Request.ServerVariables("URL"))  'o/p:- BookingForm.asp
            end if               
             
        '************************************End of this************************************************************                
                
'Here we are getting the data of the tour reservation for edit

Dim osqlReservation,Counter,RsReservation,RsTourdetails,RsExtrasDetails,osqlTour,osqlExtras,oExcCode,RsBooking,BookingId,NameonCard,strhidden,CardType
BookingId=0
if (Session("LoginUser")="")then
 Response.Redirect("checkout.asp?ReturnUrl=EditReservation")
else
oUserId=Session("LoginUser")
end if 
Open_Connection
set RsBooking=oConn.Execute("Select * from BookingInfo  where  Id=(Select max (Id) as Id from BookingInfo where UserId="&oUserId&")")
'***Check the Booking is there in the system for the Logged in user
if NOT RsBooking.EOF then
BookingId=RsBooking.Fields.Item("Id").Value

osqlReservation="Select * from TourReservation where IsActive=1 and UserId="&oUserId&" and Bookingid="&BookingId
Set RsReservation=oConn.Execute(osqlReservation)
Counter=0

 ' Check the Record set is empty 
     if NOT RsReservation.EOF then
     
        While NOT RsReservation.EOF
    
    '************Here adding this code on 06-05-2012 for custome*************    '
    if(LCase(RsReservation.Fields.Item("runtype").Value)= LCase("Custome")) then
                %>
                <!--#include file="customize/EditReservationCustome.asp"-->
                <%   
    else
    '************End of adding this code on 06-05-2012 for custome*************
    
    Counter=Counter+1
    IpAddress=RsReservation.Fields.Item("IpAddress").Value
    ReservationDate=RsReservation.Fields.Item("DefaultDate").Value
    
    oExcCode=RsReservation.Fields.Item("TourName").Value
    oNumberofpass=RsReservation.Fields.Item("NoPassenger").Value
    oReservationId=RsReservation.Fields.Item("Id").Value
    
    '***Check for the Vehicle type***
if (oNumberofpass=1 or oNumberofpass=2 ) then
Vehicle="Sedan"
end if 
if (oNumberofpass=3 or oNumberofpass=4 ) then
Vehicle="Sedan/Minivan"
end if 
if (oNumberofpass >=5 and oNumberofpass <=8 ) then
Vehicle="Sedan/Minivan"
end if 
if (oNumberofpass >=9 and oNumberofpass <=16 ) then
Vehicle="Minivan"
end if 
if (oNumberofpass >=17 and oNumberofpass <=50 ) then
Vehicle="Minibus/Bus"
end if 
'***End***
    
    '***Getting the reserved tour details***  If Counter <= iToursLength Then for the Tours and else part for the transfer service //Runtype       
    if(LCase(RsReservation.Fields.Item("Runtype").Value)<>LCase("Tour") and LCase(RsReservation.Fields.Item("Runtype").Value)<>LCase("Shore Excursion") and LCase(RsReservation.Fields.Item("Runtype").Value)<>LCase("Tour en route") and LCase(RsReservation.Fields.Item("Runtype").Value)<>LCase("Tourenroute")) then
    
    if(oExcCode="") then
    oExcCode=0
    End If
    
    set RsTourdetails = oConn.Execute("Select top 1 dbo213093820.TransferBooking.Id,TransferPricePax.Pax, MaxPax,Runtype, FromTo,Vehicle, TransferPricePax.Price,dbo213093820.TransferBooking.IsActive,Region,ImageUrl from dbo213093820.TransferBooking inner join dbo213093820.RunTypeMaster ON dbo213093820.RunTypeMaster.Id=dbo213093820.TransferBooking.RunTypeId inner join TransferPricePax ON TransferPricePax.TransferBookingId=TransferBooking.Id Inner join dbo213093820.VehicleMaster ON dbo213093820.VehicleMaster.Id=dbo213093820.TransferPricePax.VehicleId Inner join dbo213093820.FromToMaster ON dbo213093820.FromToMaster.Id=dbo213093820.TransferBooking.FromToId left outer join dbo213093820.RegionMaster ON  dbo213093820.TransferBooking.RegionId=dbo213093820.RegionMaster.Id where dbo213093820.TransferBooking.Id="&oExcCode)
    
    else
    '***Query for the tour/shore Excursion***
    set RsTourdetails = oConn.Execute("SELECT * FROM Excursions as a LEFT JOIN Tour as b ON a.DepartureCityID = b.TourID Where ExcCode ='"&oExcCode&"'")
    
    end if 
  
    if NOT RsTourdetails.EOF then
                %>
                <form name="EditReservation" action="EditBooking_Reservation.asp" method="post">
                    <br />
                    <!--Static Section -->
                    <div class="reservationcss">
                        <fieldset>
                            <%
                        
                        'if(RsReservation.Fields.Item("Runtype").Value<>"Transfer") then
                        if(LCase(RsReservation.Fields.Item("Runtype").Value)=LCase("Tour") or LCase(RsReservation.Fields.Item("Runtype").Value)=LCase("Shore Excursion") or LCase(RsReservation.Fields.Item("Runtype").Value)=LCase("Tour en route") or LCase(RsReservation.Fields.Item("Runtype").Value)=LCase("Tourenroute")) then
                        '***for the tour/shore Excursion***
                            %>
                            <legend>
                                <%=UCase(oExcCode)%>
                                &nbsp;-<%=RsTourdetails.Fields.Item("Destinations").Value  %></legend>
                            <%
                          else
                          
                            %>
                            <legend>
                                <%=RsTourdetails.Fields.Item("RunType").Value %>
                                &nbsp;-<%=RsTourdetails.Fields.Item("FromTo").Value%></legend>
                            <%End if  %>

                            <div class="formcss">
                                <table cellspacing="0" cellpadding="0" border="0" class="smalltopheadings">
                                    <tbody>
                                        <tr>
                                            <td width="280" class="paddleft">Services</td>
                                            <td width="92">Passenger#</td>
                                            <td width="64">Vehicle</td>
                                            <td>Price</td>
                                        </tr>
                                    </tbody>
                                </table>
                                <% 
                          
                          if(LCase(RsReservation.Fields.Item("Runtype").Value)=LCase("Tour") or LCase(RsReservation.Fields.Item("Runtype").Value)=LCase("Shore Excursion") or LCase(RsReservation.Fields.Item("Runtype").Value)=LCase("Tour en route") or LCase(RsReservation.Fields.Item("Runtype").Value)=LCase("Tourenroute")) then
                                %>
                                <div class="productdetails">
                                    <table cellspacing="0" cellpadding="0" border="0">
                                        <tbody>
                                            <tr>
                                                <td>
                                                    <img height="68px" width="69px" border="0" align="left" class="floatleftcss" alt=""
                                                        src="<%=RsTourdetails.Fields.Item("Image1URL").Value %>" /></td>
                                                <td>
                                                    <table width="100%" cellspacing="0" cellpadding="0" border="0">
                                                        <tbody>
                                                            <tr>
                                                                <td width="254">
                                                                    <p>
                                                                        <span class="fontboldcss">
                                                                            <%=UCase(oExcCode)%>
                                                                            &nbsp;-
                                                                            <%=RsTourdetails.Fields.Item("ActualDuration").Value%>
                                                                            Hour Tour
                                                                            <%=RsTourdetails.Fields.Item("Destinations").Value %>
                                                                        </span>
                                                                        <br />
                                                                        <span class="underlinetext">This tour start from the The
                                                                            <%=RsTourdetails.Fields.Item("TourName").Value %>
                                                                        </span>
                                                                    </p>
                                                                </td>
                                                                <td width="63" class="textcenter">
                                                                    <input type="text" id="txtPassenger_<%=LCase(oExcCode)%>" name="txtPassenger_<%=LCase(oExcCode)%>"
                                                                        onkeypress="return blockNonNumbers(this, event, false, false);" class="smallinput" readonly="readonly"
                                                                        maxlength="2" value="<%=oNumberofpass %>" />
                                                                    <input type="hidden" value="<%=oNumberofpass%>" name="hdnpassenger<%=Counter %>"
                                                                        id="hdnpassenger<%=Counter %>" />
                                                                    <input type="hidden" name="hdnTourName<%=Counter %>" id="hdnTourName<%=Counter %>"
                                                                        value="<%=UCase(oExcCode)%>" />
                                                                    <input type="hidden" id="hdnBlackoutTour<%=Counter %>" name="hdnBlackoutTour<%=Counter %>" value="<%=UCase(oExcCode)%>" />
                                                                </td>
                                                                <td width="97" class="textcenter">
                                                                    <span style="display: block; padding-left: 24px;">
                                                                        <%=Vehicle%>
                                                                    </span>
                                                                    <input type="hidden" id="hdnVehicle<%=Counter %>" name="hdnVehicle<%=Counter %>"
                                                                        value="<%=Vehicle%>" />
                                                                </td>
                                                                <td width="62" class="textcenter">
                                                                    <span class="pricetag">
                                                                        <%=RsReservation.Fields.Item("EachTourprice").Value %>
                                                                    </span>
                                                                </td>
                                                            </tr>
                                                        </tbody>
                                                    </table>
                                                    <table cellpadding="0" cellspacing="0" border="0" width="100%">
                                                        <tr>
                                                            <td width="188">
                                                                <br>
                                                                <span><strong>Total for this tour:</strong> </span>
                                                            </td>
                                                            <td width="188">&nbsp;</td>
                                                            <td width="62" class="textcenter">
                                                                <br />
                                                                <span class="pricetag">
                                                                    <%=RsReservation.Fields.Item("TourTotal").Value %>
                                                                </span>
                                                            </td>
                                                            <td>
                                                                <input type="hidden" value="<%=BookingId%>" id="hdnBookingId<%=Counter %>" name="hdnBookingId<%=Counter %>" /></td>
                                                        </tr>

                                                    </table>

                                                    <!--here including section of Supplement price add on 23-03-3012-->

                                                    <table width="100%" cellspacing="0" cellpadding="0" border="0">
                                                        <tbody>

                                                            <tr>
                                                                <td width="220">
                                                                    <span class="italiccss">Early Morning/Night Supplement: 
                                                                    </span>
                                                                </td>
                                                                <td width="138">&nbsp;</td>

                                                                <td width="62" class="textcenter">
                                                                    <span class="pricetag">
                                                                        <%=RsReservation("EachTourSupplementPrice") %>
                                                                    </span>
                                                                </td>
                                                            </tr>


                                                        </tbody>
                                                    </table>

                                                    <!--*******End of supplement section****************-->
                                                    <% 
                                                    '***Here we are getting the Extras details***
  
                                                    osqlExtras="select ExtrasPrices.ExtrasId,ShortDescExtra,ExtraPaxPrice,IsAddOn,AddOnFrom,AddOnPrice from Extras inner join ExtrasPrices on Extras.Id=ExtrasPrices.ExtrasId where ExtrasPrices.ExtrasId in ("&RsReservation.Fields.Item("ExtrasId1").Value&","&RsReservation.Fields.Item("ExtrasId2").Value &","&RsReservation.Fields.Item("ExtrasId3").Value &") and ExtraFrom <="&oNumberofpass&" and ExtraTo>="&oNumberofpass&""
  
                                                    set RsExtrasDetails=oConn.Execute(osqlExtras)
                                                    if NOT RsExtrasDetails.EOF then
                                                    %>
                                                    <h5 class="extratext">
                                                        <span class="italiccss">Extra includsdsdsed:</span></h5>
                                                    <table width="100%" cellspacing="0" cellpadding="0" border="0">
                                                        <tbody>
                                                            <% While NOT RsExtrasDetails.EOF %>
                                                            <tr>
                                                                <td width="220">
                                                                    <span class="italiccss">
                                                                        <%=RsExtrasDetails.Fields.Item("ShortDescExtra").Value%>
                                                                    </span>
                                                                </td>
                                                                <td width="138">&nbsp;</td>

                                                                <%
                                                            '***Code for the Description showing  if Addon is checked then we will show addon description and multiply the (17) No of passenger * Price for the AddOn(4)**** 
                                                            oExtrasVal=RsExtrasDetails.Fields.Item("ExtraPaxPrice").Value
                                                            if (RsExtrasDetails.Fields.Item("IsAddOn").Value="True" )then
                                                            if ( Cint(oNumberofpass)>=Cint(RsExtrasDetails.Fields.Item("AddOnFrom").Value)) then
                                                            oExtrasVal=Cint(oExtrasVal)+(Cint(oNumberofpass) * Cint(RsExtrasDetails.Fields.Item("AddOnPrice").Value))
                                                            end if
                                                            end if
        
                                                                %>

                                                                <td width="62" class="textcenter">
                                                                    <span class="pricetag">
                                                                        <%=oExtrasVal%>
                                                                    </span>
                                                                </td>
                                                            </tr>
                                                            <% 
                                                            RsExtrasDetails.MoveNext
                                                            Wend 'end RsExtrasDetails.EOF
                                                            End if 
                                                            %>
                                                        </tbody>
                                                    </table>
                                                </td>
                                            </tr>
                                        </tbody>
                                    </table>
                                </div>
                                <% Else
                                Vehicle=RsTourdetails.Fields.Item("Vehicle").Value
                                %>
                                <!--***Start HTML of the details of the Tranfer Service***-->
                                <div class="productdetails">
                                    <table cellspacing="0" cellpadding="0" border="0">
                                        <tbody>
                                            <tr>
                                                <td>
                                                    <img height="68px" width="69px" border="0" align="left" class="floatleftcss" alt=""
                                                        src="<%=RsTourdetails.Fields.Item("ImageURL").Value %>" /></td>
                                                <td>
                                                    <table width="100%" cellspacing="0" cellpadding="0" border="0">
                                                        <tbody>
                                                            <tr>
                                                                <td width="254">

                                                                    <p>
                                                                        <span class="fontboldcss">&nbsp;-
                                                                            <%=RsTourdetails.Fields.Item("Runtype").Value%>

                                                                            <%=RsTourdetails.Fields.Item("FromTo").Value %>
                                                                        </span>

                                                                    </p>

                                                                </td>
                                                                <td width="63" class="textcenter">
                                                                    <input type="text" id="Text1" name="txtPassenger_<%=LCase(oExcCode)%>"
                                                                        onkeypress="return blockNonNumbers(this, event, false, false);" class="smallinput"
                                                                        maxlength="2" value="<%=oNumberofpass %>" />
                                                                    <input type="hidden" value="<%=oNumberofpass%>" name="hdnpassenger<%=Counter %>"
                                                                        id="hdnpassenger<%=Counter %>" />
                                                                    <input type="hidden" name="hdnTourName<%=Counter %>" id="hdnTourName<%=Counter %>"
                                                                        value="<%=UCase(oExcCode)%> - <%=RsTourdetails.Fields.Item("Runtype").Value %> <%=RsTourdetails.Fields.Item("FromTo").Value %>" />
                                                                </td>
                                                                <td width="97" class="textcenter">
                                                                    <span style="display: block; padding-left: 24px;">
                                                                        <%=Vehicle%>
                                                                    </span>
                                                                    <input type="hidden" id="hdnVehicle<%=Counter %>" name="hdnVehicle<%=Counter %>"
                                                                        value="<%=Vehicle%>" />
                                                                </td>
                                                                <td width="62" class="textcenter">
                                                                    <span class="pricetag">
                                                                        <%=RsReservation.Fields.Item("EachTourprice").Value %>
                                                                    </span>
                                                                </td>
                                                            </tr>
                                                        </tbody>
                                                    </table>
                                                    <% 
                                                    '***Here we are getting the Extras details Transfer Service***
  
                                                    set RsExtraTransfer=oConn.Execute("select ExtrasId1,ExtrasId2,ExtrasId3,ExtrasId4,ExtrasId5 from TourReservation where Id="&RsReservation.Fields.Item("Id").Value&"") 
  

                                                    If NOT RsExtraTransfer.EOF  Then
                                                    %>
                                                    <h5 class="extratext">
                                                        <span class="italiccss">Extra included:</span></h5>
                                                    <%

                                                    osqlExtras="select ExtrasPrices.ExtrasId,ShortDescExtra,LongDescExtra,IsperPerson,IsperGroup,IsAddOn,AddOnPrice,AddOnFrom,ShortDescAddOn,LongDescAddOn,ExtraPax,ExtraPaxPrice,IsGroup,GroupPax,GroupPaxPrice from Extras inner join ExtrasPrices on Extras.Id=ExtrasPrices.ExtrasId where ExtrasPrices.ExtrasId IN ("&RsExtraTransfer.Fields.Item("ExtrasId1").Value&","&RsExtraTransfer.Fields.Item("ExtrasId2").Value&","&RsExtraTransfer.Fields.Item("ExtrasId3").Value&","&RsExtraTransfer.Fields.Item("ExtrasId4").Value&","&RsExtraTransfer.Fields.Item("ExtrasId5").Value&") and ExtraFrom <="&oNumberofpass&" and ExtraTo>="&oNumberofpass
                                                         
                                                    set RsExtrasDetails=oConn.Execute(osqlExtras)
                                                    if NOT RsExtrasDetails.EOF then
                                                    While NOT RsExtrasDetails.EOF 
 
                                                    %>

                                                    <table width="100%" cellspacing="0" cellpadding="0" border="0">
                                                        <tbody>

                                                            <tr>
                                                                <td width="220">
                                                                    <span class="italiccss">
                                                                        <%=RsExtrasDetails.Fields.Item("ShortDescExtra").Value%>
                                                                    </span>
                                                                </td>
                                                                <td width="128">&nbsp;</td>
                                                                <%
                                                                '***Code for the Description showing  if Addon is checked then we will show addon description and multiply the (17) No of passenger * Price for the AddOn(4)**** 
                                                                oExtrasVal=RsExtrasDetails.Fields.Item("ExtraPaxPrice").Value
                                                                if (RsExtrasDetails.Fields.Item("IsAddOn").Value="True" )then
                                                                    if ( Cint(oNumberofpass)>=Cint(RsExtrasDetails.Fields.Item("AddOnFrom").Value)) then      
                                                                         oExtrasVal=Cint(oExtrasVal)+(Cint(oNumberofpass) * Cint(RsExtrasDetails.Fields.Item("AddOnPrice").Value))
                                                                    end if
                                                                end if        
                                                                %>
                                                                <td width="62" class="textcenter">
                                                                    <span class="pricetag">
                                                                        <%=oExtrasVal%>
                                                                    </span>
                                                                </td>
                                                            </tr>
                                                            <% 
                                                                RsExtrasDetails.MoveNext
                                                                Wend
                                                                End if 'End of if NOT RsExtrasDetails.EOF then
                                                                End if 'End of if NOT RsExtrasDetails.EOF then
    
                                                            %>
                                                            <tr>
                                                                <td width="188">
                                                                    <br />
                                                                    <span><strong>Total for this service:</strong> </span>
                                                                </td>
                                                                <td width="136">&nbsp;</td>
                                                                <td width="62" class="textcenter">
                                                                    <br />
                                                                    <span class="pricetag">
                                                                        <%=RsReservation.Fields.Item("TourTotal").Value %>
                                                                    </span>
                                                                </td>
                                                                <td>
                                                                    <input type="hidden" name="hdnBookingId<%=Counter%>" id="Hidden4"
                                                                        value="<%=BookingId%>" /></td>
                                                            </tr>

                                                            <!--here including section of Supplement price add on 23-03-3012-->
                                                            <% 
                                                                'Check if notnull and price exists then only show this section
                                                            if(IsNull(RsReservation("EachTourSupplementPrice")) or RsReservation("EachTourSupplementPrice")="0") then %>
                                                        </tbody>
                                                    </table>

                                                    <% else
                                                    %>
                                        </tbody>
                                    </table>
                                    <table width="100%" cellspacing="0" cellpadding="0" border="0">
                                        <tbody>

                                            <tr>
                                                <td width="220">
                                                    <span class="italiccss">Early Morning/Night Supplement: 
                                                    </span>
                                                </td>
                                                <td width="138">&nbsp;</td>

                                                <td width="62" class="textcenter">
                                                    <span class="pricetag">
                                                        <%=RsReservation("EachTourSupplementPrice") %>
                                                    </span>
                                                </td>
                                            </tr>
                                        </tbody>
                                    </table>

                                    <% End If %>
                                    <!--*******End of supplement section****************-->
                                    </td>
                                            </tr>
                                        </tbody>
                                    </table>
                                </div>
                                <%
                                End if
                                End if 
                                'if NOT RsTourdetails.EOF then   
                                '************Here adding this code on 06-05-2012 for custome*************
                                End if  
                                'End of if(RsReservation.Fields.Item("runtype").Value="custome") then
                                '************End of adding this code on 06-05-2012 for custome*************                                        
                                %>
                                <!--***End the HTML For the transfer saervice***-->
                                <table cellspacing="0" cellpadding="0" border="0" class="formreservation">
                                    <tbody>
                                        <tr>
                                            <td>
                                                <label>
                                                    Pick-Up Date (DD/MM/YYYY):</label>
                                                <table cellpadding="0" cellspacing="0">
                                                    <tr>

                                                        <td class="datepic">
                                                            <input type="text" name="pickupdate<%=Counter%>" onkeypress="return false;" readonly="readonly"
                                                                onkeydown="return false;" onchange="return fncheckdate(<%=Counter%>);" id="pickupdate<%=Counter%>"
                                                                value="<%=RsReservation.Fields.Item("PickUpDate").Value%>" size="20" />
                                                            <br />
                                                            <span style="display: none; color: Red;" id="spnerrorpickupdate<%=Counter%>">* Pick
                                                                up date field is required</span>
                                                            <input type="hidden" name="hdnpickupdate<%=Counter%>" id="hdnpickupdate<%=Counter%>" value="" />
                                                        </td>
                                                        <td></td>
                                                    </tr>
                                                </table>
                                                <input type="hidden" name="hdnReservationId<%=Counter%>" id="hdnReservationId<%=Counter%>"
                                                    value="<%=oReservationId %>" />
                                            </td>
                                            <td align="left" class="smallsel">
                                                <label>
                                                    Pick-Up Time:</label>
                                                <table cellpadding="0" cellspacing="0">
                                                    <tr>
                                                        <td>Hour:</td>
                                                        <td>
                                                            <select size="1" name="PickUphr<%=Counter%>" id="PickUphr<%=Counter%>">
                                                                <option value="00:" <% if RsReservation.fields.item("pickuphour").value ="00:" then %>
                                                                    selected <% end if %>>00:</option>
                                                                <option value="1:" <% if RsReservation.fields.item("pickuphour").value ="1:" then %>
                                                                    selected <% end if %>>1:</option>
                                                                <option value="2:" <% if RsReservation.fields.item("pickuphour").value ="2:" then %>
                                                                    selected <% end if %>>2:</option>
                                                                <option value="3:" <% if RsReservation.fields.item("pickuphour").value ="3:" then %>
                                                                    selected <% end if %>>3:</option>
                                                                <option value="4:" <% if RsReservation.fields.item("pickuphour").value ="4:" then %>
                                                                    selected <% end if %>>4:</option>
                                                                <option value="5:" <% if RsReservation.fields.item("pickuphour").value ="5:" then %>
                                                                    selected <% end if %>>5:</option>
                                                                <option value="6:" <% if RsReservation.fields.item("pickuphour").value ="6:" then %>
                                                                    selected <% end if %>>6:</option>
                                                                <option value="7:" <% if RsReservation.fields.item("pickuphour").value ="7:" then %>
                                                                    selected <% end if %>>7:</option>
                                                                <option value="8:" <% if RsReservation.fields.item("pickuphour").value ="8:" then %>
                                                                    selected <% end if %>>8:</option>
                                                                <option value="9:" <% if RsReservation.fields.item("pickuphour").value ="9:" then %>
                                                                    selected <% end if %>>9:</option>
                                                                <option value="10:" <% if RsReservation.fields.item("pickuphour").value ="10:" then %>
                                                                    selected <% end if %>>10:</option>
                                                                <option value="11:" <% if RsReservation.fields.item("pickuphour").value ="11:" then %>
                                                                    selected <% end if %>>11:</option>
                                                                <option value="12:" <% if RsReservation.fields.item("pickuphour").value ="12:" then %>
                                                                    selected <% end if %>>12:</option>
                                                                <option value="13:" <% if RsReservation.fields.item("pickuphour").value ="13:" then %>
                                                                    selected <% end if %>>13:</option>
                                                                <option value="14:" <% if RsReservation.fields.item("pickuphour").value ="14:" then %>
                                                                    selected <% end if %>>14:</option>
                                                                <option value="15:" <% if RsReservation.fields.item("pickuphour").value ="15:" then %>
                                                                    selected <% end if %>>15:</option>
                                                                <option value="16:" <% if RsReservation.fields.item("pickuphour").value ="16:" then %>
                                                                    selected <% end if %>>16:</option>
                                                                <option value="17:" <% if RsReservation.fields.item("pickuphour").value ="17:" then %>
                                                                    selected <% end if %>>17:</option>
                                                                <option value="18:" <% if RsReservation.fields.item("pickuphour").value ="18:" then %>
                                                                    selected <% end if %>>18:</option>
                                                                <option value="19:" <% if RsReservation.fields.item("pickuphour").value ="19:" then %>
                                                                    selected <% end if %>>19:</option>
                                                                <option value="20:" <% if RsReservation.fields.item("pickuphour").value ="20:" then %>
                                                                    selected <% end if %>>20:</option>
                                                                <option value="21:" <% if RsReservation.fields.item("pickuphour").value ="21:" then %>
                                                                    selected <% end if %>>21:</option>
                                                                <option value="22:" <% if RsReservation.fields.item("pickuphour").value ="22:" then %>
                                                                    selected <% end if %>>22:</option>
                                                                <option value="23:" <% if RsReservation.fields.item("pickuphour").value ="23:" then %>
                                                                    selected <% end if %>>23:</option>
                                                                <option value="24:" <% if RsReservation.fields.item("pickuphour").value ="24:" then %>
                                                                    selected <% end if %>>24:</option>
                                                            </select>
                                                        </td>
                                                        <td>Minute:</td>
                                                        <td>
                                                            <select size="1" name="PickUpMin<%=Counter%>" id="PickUpMin<%=Counter%>">

                                                                <option value="00" <% if RsReservation.fields.item("pickupmin").value ="00" then %>
                                                                    selected <% end if %>>00</option>
                                                                <option value="05" <% if RsReservation.fields.item("pickupmin").value ="05" then %>
                                                                    selected <% end if %>>05</option>
                                                                <option value="10" <% if RsReservation.fields.item("pickupmin").value ="10" then %>
                                                                    selected <% end if %>>10</option>
                                                                <option value="10" <% if RsReservation.fields.item("pickupmin").value ="15" then %>
                                                                    selected <% end if %>>15</option>
                                                                <option value="20" <% if RsReservation.fields.item("pickupmin").value ="20" then %>
                                                                    selected <% end if %>>20</option>
                                                                <option value="25" <% if RsReservation.fields.item("pickupmin").value ="25" then %>
                                                                    selected <% end if %>>25</option>
                                                                <option value="30" <% if RsReservation.fields.item("pickupmin").value ="30" then %>
                                                                    selected <% end if %>>30</option>
                                                                <option value="35" <% if RsReservation.fields.item("pickupmin").value ="35" then %>
                                                                    selected <% end if %>>35</option>
                                                                <option value="40" <% if RsReservation.fields.item("pickupmin").value ="40" then %>
                                                                    selected <% end if %>>40</option>
                                                                <option value="45" <% if RsReservation.fields.item("pickupmin").value ="45" then %>
                                                                    selected <% end if %>>45</option>
                                                                <option value="50" <% if RsReservation.fields.item("pickupmin").value ="50" then %>
                                                                    selected <% end if %>>50</option>
                                                                <option value="55" <% if RsReservation.fields.item("pickupmin").value ="55" then %>
                                                                    selected <% end if %>>55</option>
                                                            </select>
                                                        </td>
                                                    </tr>
                                                </table>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td align="left" colspan="2">
                                                <label>
                                                    Passenger Leading Name:</label><input type="text" name="LeadName<%=Counter%>" id="LeadName<%=Counter%>"
                                                        value="<%=RsReservation.Fields.Item("LeadingName").Value%>" title="This name will be used on the greeting sign"
                                                        size="20" /></td>
                                        </tr>
                                        <tr>
                                            <td colspan="2">
                                                <table>
                                                    <tbody>
                                                        <tr>
                                                            <td align="left">
                                                                <label>
                                                                    Total No of people travelling:</label></td>
                                                            <td>
                                                                <input type="text" width="16" name="txtNoTrave<%=Counter%>" id="txtNoTrave<%=Counter%>"
                                                                    disabled="disabled" value="<%=RsReservation.Fields.Item("NoPassenger").Value%>"
                                                                    class="smallinput" maxlength="2" title="people travelling" /></td>
                                                            <td>
                                                                <label>
                                                                    Adult:</label></td>
                                                            <td>
                                                                <input type="text" width="16" name="txtAdult<%=Counter%>" id="txtAdult<%=Counter%>"
                                                                    onchange="fnFillChildren(<%=Counter%>);" value="<%=RsReservation.Fields.Item("NoAdult").Value%>"
                                                                    class="smallinput" maxlength="2" onkeypress="return blockNonNumbers(this, event, false, false);" /></td>
                                                            <td>
                                                                <label>
                                                                    Children under the age of 12:</label></td>
                                                            <td>
                                                                <input type="text" width="16" id="txtChildren<%=Counter%>" onchange="fnFillAdult(<%=Counter%>);"
                                                                    name="txtChildren<%=Counter%>" class="smallinput" onkeypress="return blockNonNumbers(this, event, false, false);"
                                                                    maxlength="2" value="<%=RsReservation.Fields.Item("NoChildren").Value%>" /></td>
                                                        </tr>
                                                    </tbody>
                                                </table>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td width="249" align="left">
                                                <label>
                                                    Pick-Up From:</label>
                                                <select onchange="ChangePickupfields(<%=Counter%>)" style="color: rgb(160, 111, 83); font-size: 10px; position: relative; width: 147px;"
                                                    id="ddlpickup<%=Counter%>"
                                                    name="ddlpickup<%=Counter%>">

                                                    <option value="Airport" <% if RsReservation.fields.item("PickFrom").value ="Airport" then %>selected <% end if %>>Airport</option>
                                                    <option value="Hotel" <% if RsReservation.fields.item("PickFrom").value ="Hotel" then %>selected <% end if %>>Hotel</option>
                                                    <option value="Other" <% if RsReservation.fields.item("PickFrom").value ="Other" then %>selected <% end if %>>Other</option>
                                                    <option value="Port" <% if RsReservation.fields.item("PickFrom").value ="Port" then %>selected <% end if %>>Port</option>
                                                    <option value="Train" <% if RsReservation.fields.item("PickFrom").value ="Train" then %>selected <% end if %>>Train Station</option>
                                                    <option value="Villa" <% if RsReservation.fields.item("PickFrom").value ="Villa" then %>selected <% end if %>>Vill/Apartment</option>
                                                </select>

                                            </td>
                                        </tr>
                                        <tr>
                                            <td colspan="2" id="tdpickupfields1">
                                                <%  
                                                If RsReservation.Fields.Item("PickFrom").Value ="Airport" Then
                                                style="style='display:block;'"
                                                else
                                                style="style='display:none;'"
                                                end if 
                                                %>
                                                <table cellpadding="0" cellspacing="0" id="tblpickupAirport<%=Counter%>">
                                                    <tbody>
                                                        <tr>
                                                            <td align="left">
                                                                <label>
                                                                    Which airport/city will your flight arrive into:</label>
                                                                <input type="text" name="arrivalcity<%=Counter%>" id="arrivalcity<%=Counter%>" value="<%=RsReservation.Fields.Item("PickUpCityFlightArrive").Value%>"
                                                                    size="20" />
                                                            </td>
                                                            <td width="249" align="left">
                                                                <label>
                                                                    Flight Number:</label>
                                                                <input type="text" name="flightnumber<%=Counter%>" id="flightnumber<%=Counter%>"
                                                                    value="<%=RsReservation.Fields.Item("PickUpFlightNo").Value%>" size="20" />
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td align="left">
                                                                <label>
                                                                    Airline</label>
                                                                <input type="text" name="airline<%=Counter%>" id="airline<%=Counter%>" value="<%=RsReservation.Fields.Item("PickupAirline").Value%>"
                                                                    size="20" />
                                                            </td>
                                                        </tr>
                                                    </tbody>
                                                </table>
                                                <%
         
                                                If RsReservation.Fields.Item("PickFrom").Value ="Train" Then
                                                style="style='display:block;'"
                                                else
                                                style="style='display:none;'"
                                                end if
                                                %>
                                                <table cellpadding="0" cellspacing="0" id="tblpickupTrain<%=Counter%>">
                                                    <tbody>
                                                        <tr>
                                                            <td align="left">
                                                                <label>
                                                                    Which station/city will your train arrive:</label>
                                                                <input type="text" name="Trainarrivalcity<%=Counter%>" id="Trainarrivalcity<%=Counter%>"
                                                                    value="<%=RsReservation.Fields.Item("PickUpCityTrainArrive").Value%>" size="20" />
                                                            </td>
                                                            <td width="249" align="left">
                                                                <label>
                                                                    Train Number:</label>
                                                                <input type="text" name="Trainnumber<%=Counter%>" id="Trainnumber<%=Counter%>" value="<%=RsReservation.Fields.Item("PickUpTrainNo").Value%>"
                                                                    size="20" />
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td align="left">
                                                                <label>
                                                                    Train Departing From:</label>
                                                                <input type="text" name="TrainDepart<%=Counter%>" id="TrainDepart<%=Counter%>" size="20"
                                                                    value="<%=RsReservation.Fields.Item("PickUpTrainDepartFrom").Value%>" title="Train Departing From" />
                                                            </td>
                                                            <% '*****Adding Train name field for pickup %>
                                                            <td width="249" align="left">
                                                                <label>
                                                                    Train Name:</label>
                                                                <input type="text" name="Trainname<%=Counter%>" id="Trainname<%=Counter%>" value="<%=RsReservation.Fields.Item("PickUpArriveTrainName").Value%>"
                                                                    size="20" />
                                                            </td>
                                                            <% '*****End of Adding Train name field for pickup %>
                                                        </tr>
                                                    </tbody>
                                                </table>
                                                <%
        
                                                If RsReservation.Fields.Item("PickFrom").Value ="Port" Then
                                                style="style='display:block;'"
                                                else
                                                style="style='display:none;'"
                                                end if
                                                %>
                                                <table cellpadding="0" cellspacing="0" id="tblpickupPort<%=Counter%>">
                                                    <tbody>
                                                        <tr>
                                                            <td align="left">
                                                                <label>
                                                                    Which port will your boat dock:</label>
                                                                <input type="text" name="arrivalport<%=Counter%>" id="arrivalport<%=Counter%>" value="<%=RsReservation.Fields.Item("PickUpPort").Value%>"
                                                                    title="Port will your boat dock" size="20" />
                                                            </td>
                                                            <td width="249" align="left">
                                                                <label>
                                                                    Ship/Cruise Name:</label>
                                                                <input type="text" name="shipname<%=Counter%>" id="shipname<%=Counter%>" value="<%=RsReservation.Fields.Item("PickupShipName").Value%>"
                                                                    size="20" title="Ship/Cruise Name" />
                                                            </td>
                                                        </tr>
                                                    </tbody>
                                                </table>
                                                <%
        
                                                If RsReservation.Fields.Item("PickFrom").Value ="Hotel" Then
                                                style="style='display:block;'"
                                                else
                                                style="style='display:none;'"
                                                end if
                                                %>
                                                <table cellpadding="0" cellspacing="0" id="tblpickupHotel<%=Counter%>">
                                                    <tbody>
                                                        <tr>
                                                            <td align="left">
                                                                <label>
                                                                    Hotel Name:</label>
                                                                <input type="text" name="hotelname<%=Counter%>" id="hotelname<%=Counter%>" value="<%=RsReservation.Fields.Item("PickUpHotelName").Value%>"
                                                                    title="Hotel Name" size="20" />
                                                            </td>
                                                            <td width="249" align="left">
                                                                <label>
                                                                    City:</label>
                                                                <input type="text" id="hotelcity<%=Counter%>" name="hotelcity<%=Counter%>" title="Hotel City"
                                                                    value="<%=RsReservation.Fields.Item("PickUpHotelCity").Value%>" size="20" />
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td align="left">
                                                                <label>
                                                                    Hotel Address (if available):</label>
                                                                <input type="text" id="hoteladdress<%=Counter%>" name="hoteladdress<%=Counter%>"
                                                                    value="<%=RsReservation.Fields.Item("PickUpHotelAddress").Value%>" title="Hotel Address" />
                                                            </td>
                                                        </tr>
                                                    </tbody>
                                                </table>
                                                <!--Table for pick up Villa-->
                                                <table cellpadding="0" cellspacing="0" id="tblpickupVilla<%=Counter%>">
                                                    <tbody>
                                                        <tr>
                                                            <td align="left">
                                                                <label>
                                                                    Name of the Villa/Apartment:</label>
                                                                <input type="text" name="nameofvilla<%=Counter%>" id="nameofvilla<%=Counter%>" value="<%=RsReservation.Fields.Item("Pickupnameofvilla").Value%>"
                                                                    title="Name of Villa" size="20" />
                                                            </td>
                                                            <td width="249" align="left">
                                                                <label>
                                                                    Address::</label>
                                                                <input type="text" id="villaAddress<%=Counter%>" name="villaAddress<%=Counter%>"
                                                                    title="Address" value="<%=RsReservation.Fields.Item("PickupvillaAddress").Value%>"
                                                                    size="20" />
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td align="left">
                                                                <label>
                                                                    City:</label>
                                                                <input type="text" id="villaCity<%=Counter%>" name="villaCity<%=Counter%>" value="<%=RsReservation.Fields.Item("PickupvillaCity").Value%>"
                                                                    title="Hotel Address" />
                                                            </td>
                                                        </tr>
                                                    </tbody>
                                                </table>
                                                <!--Table for Pick up Other-->
                                                <table cellpadding="0" cellspacing="0" id="tblpickupOther<%=Counter%>">
                                                    <tbody>
                                                        <tr>
                                                            <td align="left">
                                                                <label>
                                                                    Other :</label>
                                                                <textarea cols="16" rows="5" id="OtherPickup<%=Counter%>" name="OtherPickup<%=Counter%>"><%=RsReservation.Fields.Item("PickUpOther").Value%></textarea>
                                                            </td>
                                                        </tr>
                                                    </tbody>
                                                </table>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td align="left">
                                                <label>
                                                    Drop-off:</label>
                                                <select onchange="ChangeDropofffields(<%=Counter%>);" style="color: rgb(160, 111, 83); font-size: 10px; position: relative; width: 147px;"
                                                    id="ddlDropOff<%=Counter%>"
                                                    name="ddlDropOff<%=Counter%>">

                                                    <option value="Airport" <% if RsReservation.fields.item("dropoff").value ="Airport" then %>
                                                        selected <% end if %>>Airport</option>
                                                    <option value="Hotel" <% if RsReservation.fields.item("dropoff").value ="Hotel" then %>
                                                        selected <% end if %>>Hotel</option>
                                                    <option value="Other" <% if RsReservation.fields.item("dropoff").value ="Other" then %>
                                                        selected <% end if %>>Other</option>
                                                    <option value="Port" <% if RsReservation.fields.item("dropoff").value ="Port" then %>
                                                        selected <% end if %>>Port</option>
                                                    <option value="Train" <% if RsReservation.fields.item("dropoff").value ="Train" then %>
                                                        selected <% end if %>>Train Station</option>
                                                    <option value="Villa" <% if RsReservation.fields.item("dropoff").value ="Villa" then %>
                                                        selected <% end if %>>Villa/Apartment</option>
                                                </select>
                                                <!--<option value="Other">Other</option>-->
                                            </td>
                                            <td align="left"></td>
                                        </tr>
                                        <%  
                                        If RsReservation.Fields.Item("DropOff").Value ="Airport" Then
                                        style="style='display:block;'"
                                        else
                                        style="style='display:none;'"
                                        end if 
                                        %>
                                        <tr>
                                            <td id="tdshowpickdrop<%=Counter%>" colspan="2">
                                                <table cellspacing="0" cellpadding="0" border="0" id="tbldropoffAirport<%=Counter%>">
                                                    <tbody>
                                                        <tr>
                                                            <td width="249" align="left">
                                                                <label>
                                                                    Which airport/city will your flight depart from (Drop off):</label>
                                                                <input type="text" name="arrivalcityDropoff<%=Counter%>" id="arrivalcityDropoff<%=Counter%>"
                                                                    value="<%=RsReservation.Fields.Item("DropCityFlightArrive").Value%>" size="20" /></td>
                                                            <td width="249" align="left" class="valignbot">
                                                                <label>
                                                                    Flight Number(Drop off):</label>
                                                                <input type="text" name="flightnumberDropoff<%=Counter%>" id="flightnumberDropoff<%=Counter%>"
                                                                    value="<%=RsReservation.Fields.Item("DropFlightNo").Value%>" size="20" /></td>
                                                        </tr>
                                                        <tr>
                                                            <td width="249" align="left">
                                                                <label>
                                                                    Airline(Drop off):</label>
                                                                <input type="text" name="airlineDropoff<%=Counter%>" id="airlineDropoff<%=Counter%>"
                                                                    size="20" value="<%=RsReservation.Fields.Item("DropAirline").Value%>" /></td>
                                                            <td></td>
                                                        </tr>
                                                    </tbody>
                                                </table>
                                                <%
      
                                                If RsReservation.Fields.Item("DropOff").Value ="Train" Then
                                                style="style='display:block;'"
                                                else
                                                style="style='display:none;'"
                                                end if 
                                                %>
                                                <table cellspacing="0" cellpadding="0" border="0" id="tbldropoffTrain<%=Counter%>">
                                                    <tbody>
                                                        <tr>
                                                            <td width="249" align="left">
                                                                <label>
                                                                    Which station/city will your train depart from (Drop off):</label>
                                                                <input type="text" id="TrainarrivalcityDropoff<%=Counter%>" title="city will your train arrive"
                                                                    name="TrainarrivalcityDropoff<%=Counter%>" value="<%=RsReservation.Fields.Item("DropCityTrainArrive").Value%>"
                                                                    size="20" /></td>
                                                            <td width="249" align="left" class="valignbot">
                                                                <label>
                                                                    Train Number(Drop off):</label>
                                                                <input type="text" id="TrainnumberDropoff<%=Counter%>" name="TrainnumberDropoff<%=Counter%>"
                                                                    title="Train Number" size="20" value="<%=RsReservation.Fields.Item("DropTrainNo").Value%>" /></td>
                                                        </tr>
                                                        <tr>
                                                            <td width="249" align="left">
                                                                <label>
                                                                    Train Departing From (Drop off):</label>
                                                                <input type="text" id="TrainDropoff<%=Counter%>" name="TrainDropoff<%=Counter%>"
                                                                    title="Train Departing From" size="20" value="<%=RsReservation.Fields.Item("DropTrainDepart").Value%>" /></td>


                                                            <% '*****Adding Train name field for pickup %>
                                                            <td width="249" align="left">
                                                                <label>
                                                                    Train Name:</label>
                                                                <input type="text" name="TrainnameDropoff<%=Counter%>" id="TrainnameDropoff<%=Counter%>" value="<%=RsReservation.Fields.Item("DropArriveTrainname").Value%>"
                                                                    size="20" />
                                                            </td>
                                                            <% '*****End of Adding Train name field for pickup %>
                                                        </tr>
                                                    </tbody>
                                                </table>
                                                <%       
                                                If RsReservation.Fields.Item("DropOff").Value ="Port" Then
                                                style="style='display:block;'"
                                                else
                                                style="style='display:none;'"
                                                end if 
                                                %>
                                                <table cellspacing="0" cellpadding="0" border="0" id="tbldropoffPort<%=Counter%>">
                                                    <tbody>
                                                        <tr>
                                                            <td width="249" align="left">
                                                                <label>
                                                                    Which port will your boat dock (Drop off):</label>
                                                                <input type="text" id="arrivalportDropoff<%=Counter%>" name="arrivalportDropoff<%=Counter%>"
                                                                    value="<%=RsReservation.Fields.Item("DropPort").Value%>" title="Which port will your boat dock"
                                                                    size="20" /></td>
                                                            <td width="249" align="left" class="valignbot">
                                                                <label>
                                                                    Ship/Cruise Name (Drop off):</label>
                                                                <input type="text" id="shipnameDropoff<%=Counter%>" name="shipnameDropoff<%=Counter%>"
                                                                    value="<%=RsReservation.Fields.Item("DropshipName").Value%>" title="Ship/Cruise Name"
                                                                    size="20" /></td>
                                                        </tr>
                                                    </tbody>
                                                </table>
                                                <%
    
                                                    If RsReservation.Fields.Item("DropOff").Value ="Hotel" Then
                                                    style="style='display:block;'"
                                                    else
                                                    style="style='display:none;'"
                                                    end if 
                                                %>
                                                <table cellspacing="0" cellpadding="0" border="0" id="tbldropoffHotel<%=Counter%>">
                                                    <tbody>
                                                        <tr>
                                                            <td width="249" align="left">
                                                                <label>
                                                                    Hotel Name (Drop off):</label>
                                                                <input type="text" id="hotelnameDropoff<%=Counter%>" title="Hotel Name" name="hotelnameDropoff<%=Counter%>"
                                                                    value="<%=RsReservation.Fields.Item("DropHotelName").Value%>" size="20" /></td>
                                                            <td width="249" align="left" class="valignbot">
                                                                <label>
                                                                    City (Drop off):</label>
                                                                <input type="text" id="hotelcityDropoff<%=Counter%>" name="hotelcityDropoff<%=Counter%>"
                                                                    value="<%=RsReservation.Fields.Item("DropHotelCity").Value%>" title="Hotel City"
                                                                    size="20" /></td>
                                                        </tr>
                                                        <tr>
                                                            <td width="249" align="left">
                                                                <label>
                                                                    Hotel Address (if available):</label>
                                                                <input type="text" id="hoteladdressDropoff<%=Counter%>" name="hoteladdressDropoff<%=Counter%>"
                                                                    value="<%=RsReservation.Fields.Item("DropHotelAddress").Value%>" title="Hotel Address"
                                                                    size="20" /></td>
                                                            <td></td>
                                                        </tr>
                                                    </tbody>
                                                </table>
                                                <!--Table for dropoff Villa-->
                                                <table cellpadding="0" cellspacing="0" id="tbldropoffVilla<%=Counter%>">
                                                    <tbody>
                                                        <tr>
                                                            <td align="left">
                                                                <label>
                                                                    Name of the Villa/Apartment (Drop off):</label>
                                                                <input type="text" name="nameofvillaDropoff<%=Counter%>" id="nameofvillaDropoff<%=Counter%>"
                                                                    value="<%=RsReservation.Fields.Item("Dropoffnameofvilla").Value%>" title="Name of Villa"
                                                                    size="20" />
                                                            </td>
                                                            <td width="249" align="left">
                                                                <label>
                                                                    Address (Drop off):</label>
                                                                <input type="text" id="villaAddressDropoff<%=Counter%>" name="villaAddressDropoff<%=Counter%>"
                                                                    title="Address" value="<%=RsReservation.Fields.Item("DropoffvillaAddress").Value%>"
                                                                    size="20" />
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td align="left">
                                                                <label>
                                                                    City (Drop off):</label>
                                                                <input type="text" id="villaCityDropoff<%=Counter%>" name="villaCityDropoff<%=Counter%>"
                                                                    value="<%=RsReservation.Fields.Item("DropoffvillaCity").Value%>" title="Hotel Address" />
                                                            </td>
                                                        </tr>
                                                    </tbody>
                                                </table>
                                                <!--Table for dropoff Other-->
                                                <table cellpadding="0" cellspacing="0" id="tbldropoffOther<%=Counter%>">
                                                    <tbody>
                                                        <tr>
                                                            <td align="left">
                                                                <label>
                                                                    Other (Drop off):</label>
                                                                <textarea cols="16" rows="5" id="OtherDropoff<%=Counter%>" name="OtherDropoff<%=Counter%>"><%=RsReservation.Fields.Item("DropOther").Value%></textarea>
                                                            </td>
                                                        </tr>
                                                    </tbody>
                                                </table>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td colspan="2">
                                                <label>
                                                    Notes to the booking agent:</label>
                                                <textarea name="notebookingagent<%=Counter%>" id="notebookingagent<%=Counter%>" rows="5"
                                                    cols="16"><%=RsReservation.Fields.Item("Notes").Value%></textarea>
                                            </td>
                                        </tr>
                                    </tbody>
                                </table>
                            </div>
                        </fieldset>

                        <script type="text/javascript"> 
                            $(document).ready(function () {
                                $('#pickupdate<%=Counter%>').datepicker();

                            });
                        </script>

                    </div>
                    <%                    
                   
                        RsReservation.MoveNext
                        Wend 'While NOT RsReservation.EOF
                        RsReservation.MoveFirst
                        End if ' End if NOT RsReservation.EOF then     
                        if NOT RsReservation.EOF then
                    %>
                    <!--Static Section End -->
                    <table>
                        <tr>
                            <td colspan="3" id="tdshowpickdrop"></td>
                        </tr>
                        <tr style="display: none;">
                            <td align="left" width="200">Please send confirmation for this service via:</td>
                            <td align="left" width="165">
                                <table width="75%" border="0">
                                    <tr>
                                        <td width="19%">
                                            <input type="radio" name="confirmation" id="IsEmail" <% if RsReservation.fields.item("confirmationservice").value ="Email" then %>
                                                checked <% end if %> value="Email" /></td>
                                        <td width="81%">
                                            <strong>Email</strong></td>
                                        <td>
                                            <input type="radio" name="confirmation" id="IsPhone" <% if RsReservation.fields.item("confirmationservice").value="Phone" then %>
                                                checked <% end if %> value="Phone" /></td>
                                        <td>
                                            <strong>Phone</strong></td>
                                        <td>
                                            <input type="radio" name="confirmation" id="IsFax" <% if RsReservation.fields.item("confirmationservice").value="Fax" then %>
                                                checked <% end if %> value="Fax" /></td>
                                        <td>
                                            <strong>Fax</strong></td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr style="display: none;">
                            <td align="left">
                                <i>If Phone or E-mail differs from previous details, please enter appropriately:</i></td>
                            <td align="left">
                                <input type="text" size="35" name="contactdiffers" id="contactdiffers" value="<%=RsReservation.Fields.Item("Contactdiffers").Value%>" /></td>
                        </tr>
                        <tr>
                            <td align="left" width="200">
                                <strong>How did you hear about us?:</strong></td>
                            <td align="left" width="165">
                                <select name="referral" id="referral" style="color: #a06f53; font-size: 10px; position: relative; left: 0px; width: 147px;"
                                    onchange="fncheckOther();">
                                    <option value="-------" selected>-------------------------------------</option>
                                    <option value="10 Best of Everything" <% if RsReservation.fields.item("referral").value ="10 Best of Everything" then %>
                                        selected <% end if %>>10 Best of Everything</option>
                                    <option value="Cruise Critic" <% if RsReservation.fields.item("referral").value ="Cruise Critic" then %>
                                        selected <% end if %>>Cruise Critic</option>
                                    <option value="Fodor's" <% if RsReservation.fields.item("referral").value ="Fodors" then %>
                                        selected <% end if %>>Fodor's</option>
                                    <option value="Friends" <% if RsReservation.fields.item("referral").value ="Friends" then %>
                                        selected <% end if %>>Friends</option>
                                    <option value="Frommers" <% if RsReservation.fields.item("referral").value ="Frommers" then %>
                                        selected <% end if %>>Frommer's</option>
                                    <option value="I am a repeat client" <% if RsReservation.fields.item("referral").value ="I am a repeat client" then %>
                                        selected <% end if %>>I am a repeat client</option>
                                    <option value="Internet Search" <% if RsReservation.fields.item("referral").value ="Internet Search" then %>
                                        selected <% end if %>>Internet Search</option>
                                    <option value="Limos.com" <% if RsReservation.fields.item("referral").value ="Limos.com" then %>
                                        selected <% end if %>>Limos.com</option>
                                    <option value="NLA" <% if RsReservation.fields.item("referral").value ="NLA" then %>
                                        selected <% end if %>>NLA</option>
                                    <option value="Other/Not in the list" <% if RsReservation.fields.item("referral").value ="Other" then %>
                                        selected <% end if %>>Other/Not in the list</option>
                                    <option value="Previous Client" <% if RsReservation.fields.item("referral").value ="Previous Client" then %>
                                        selected <% end if %>>Previous Client</option>
                                    <option value="Rick Steves" <% if RsReservation.fields.item("referral").value ="Rick Steves" then %>
                                        selected <% end if %>>Rick Steves</option>
                                    <option value="Slow Travel" <% if RsReservation.fields.item("referral").value ="Slow Travel" then %>
                                        selected <% end if %>>Slow Travel</option>
                                    <option value="Tour Operator" <% if RsReservation.fields.item("referral").value ="Tour Operator" then %>
                                        selected <% end if %>>Tour Operator</option>
                                    <option value="Travel Agent" <% if RsReservation.fields.item("referral").value ="Travel Agent" then %>
                                        selected <% end if %>>Travel Agent</option>
                                    <option value="Trade Show" <% if RsReservation.fields.item("referral").value ="Trade Show" then %>
                                        selected <% end if %>>Trade Show</option>
                                    <option value="Trip Advisor" <% if RsReservation.fields.item("referral").value ="Trip Advisor" then %>
                                        selected <% end if %>>Trip Advisor</option>
                                    <option value="Website" <% if RsReservation.fields.item("referral").value ="Website" then %>
                                        selected <% end if %>>Website</option>
                                </select>
                                <br />
                                <span id="spnerrorddlreferral" style="display: none; color: Red;">* Referral field is required</span>
                            </td>
                        </tr>

                        <tr id="trRefother" style="display: none;">
                            <td align="left">
                                <br />
                                <br />
                                Other:</td>
                            <td align="left">
                                <br />
                                <br />
                                <input type="text" size="20" value="<%=RsReservation.Fields.Item("Otheraboutus").Value%>"
                                    name="txtRefother" id="txtRefother" /></td>
                        </tr>
                        <tr>
                            <td align="left" width="200">Additional Information:</td>
                            <td align="left" width="165">
                                <textarea name="additionalinfo" id="additionalinfo" rows="6" cols="20"><%=RsReservation.Fields.Item("Additionalinfo").Value%></textarea></td>
                        </tr>
                        <tr>
                            <td></td>

                            <td align="left" colspan="2" width="250px;">
                                <br />
                                <input type="button" style="margin-left: -15px;" value=" " name="btnok" onmouseout="fnCorrectout();" onmouseover="fnCorrectover();" class="bgrnd_InputCorrectout" id="btnok" onclick="GotoBooking();" />
                                <input type="submit" name="btnedit" id="btnedit" onmouseout="fnEditout();" onmouseover="fnEditover();" class="bgrnd_InputEditout" value=" " onclick="return fncheckdateAll();" /><br />
                                <br />
                            </td>
                        </tr>
                    </table>
                    <%
                        '***Here End of the Reservaton Is empty checking 
                        Else
                        Response.Write("<br/> <br/><strong>There is no booking in the system for the looged in User.</strong> <br/> <br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/>")
                        MessagetoSave="No Booking1"
                        End if  
                        '***Here End of the RsBooking Is empty checking 
                        Else
                        Response.Write("<br/> <br/><strong>There is no booking in the system for the looged in User.</strong> <br/> <br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/>")
                        MessagetoSave="No Booking2"
                        End if  
                        '***Close connection
                        Close_Connection
                        '***Total of the tours Transfer Service+Tours***iTourCount is save the only tour count not include the transfer tour***

                        iTourCount=Cint(iToursLength)
                        if (Cint(itransferLength) > 0)then
                        iToursLength=Cint(iToursLength)+Cint(itransferLength)
                        end if 
           
                        '*************adding this 16-06-2012 to add custometour length also**********
                        if (Cint(iCustomeLength) > 0)then
                        iToursLength=Cint(iToursLength)+CInt(iCustomeLength)
                        end if
                        '*************End of adding this 16-06-2012 to add custometour length also**********
           
                    %>
                    <input type="hidden" id="hdnFormName" name="hdnFormName" value="EditReservation" />
                    <input type="hidden" id="hdnCounter" name="hdnCounter" value="<%=Counter%>" />
                    <input type="hidden" id="hdnIpAddress" name="hdnIpAddress" value="<%=IpAddress%>" />
                    <input type="hidden" id="hdnReservationDate" name="hdnReservationDate" value="<%=ReservationDate%>" />
                    <input type="hidden" id="hdncartcounter" name="hdncartcounter" value="<%=iToursLength%>" />
                    <input type="hidden" id="hdnTourCount" name="hdnTourCount" value="<%=iTourCount%>" />
                    <input type="hidden" id="hdnLoggedin" name="hdnLoggedin" value="<%=Session("LoginUser")%>" />
                    <input type="hidden" id="hdnError" name="hdnError" value="0" />
                </form>
            </div>
        </div>
    </div>
    <div id="MMMenuContainer0804204214_0">
        <div id="MMMenu0804204214_0" onmouseout="MM_menuStartTimeout(100);" onmouseover="MM_menuResetTimeout();">
            <a href="http://www.xyz.com/guestbook.asp" id="MMMenu0804204214_0_Item_0"
                class="MMMIFVStyleMMMenu0804204214_0" onmouseover="MM_menuOverMenuItem('MMMenu0804204214_0');">GUESTBOOK </a><a href="http://www.xyz.com/photo_gallery.asp" id="MMMenu0804204214_0_Item_1"
                    class="MMMIVStyleMMMenu0804204214_0" onmouseover="MM_menuOverMenuItem('MMMenu0804204214_0');">PHOTO&nbsp;GALLERY </a>
            <a href="Videogallery.asp" id="MMMenu0804204214_0_Item_2" class="MMMIVStyleMMMenu0804204214_0" onmouseover="MM_menuOverMenuItem('MMMenu0804204214_0');">VIDEO&nbsp;GALLERY</a>
        </div>
        <script type="text/javascript">

            ShowPickupTable();
            ShowDropoffTable();
            if (document.getElementById('hdnCounter').value > 0) {
                alert("Please check the entered details and Edit if not correct!")
            }
            if (document.getElementById("hdncartcounter").value > 0) {
                document.getElementById("spnCartCounter").innerHTML = document.getElementById("hdncartcounter").value;
                document.getElementById("spnCartCounter").className = 'cardcounter';
            }
            fncheckOther();



        </script>

    </div>
    <div id="MMMenuContainer0804204214_tourenroute">
        <div id="MMMenu0804204214_tourenroute" onmouseout="MM_menuStartTimeout(100);" onmouseover="MM_menuResetTimeout();">
            <a href="tourenroute.asp" class="MMMIFVStyleMMMenu0804204214_tourenroute" onmouseover="MM_menuOverMenuItem('MMMenu0804204214_tourenroute');">TOUR ENROUTE </a><a href="transfers.asp" onmouseover="MM_menuOverMenuItem('MMMenu0804204214_tourenroute');">TRANSFERS A TO B</a>
        </div>

        <!--//*******************Code to get Browser name 10-04-2012*****************************************-->
        <script type="text/javascript" src="js/getOtherInfo.js"></script>
        <!--//**************************************************************-->

        <script type="text/javascript">
            //Call function to hide the Sign menu option

            if (document.getElementById("hdnLoggedin").value > 0) {
                //   fnhidesign(); 
                //****here we are setting for the top your cart section***
                document.getElementById('minibag').style.left = '-122px';
                document.getElementById('spnCartCounter').style.left = '851px';

                document.getElementById("MMMenu0804204214_tourenroute").style.left = '193px';
                document.getElementById("MMMenuContainer0804204214_0").style.left = '690px';
                document.getElementById("imgleftmenu").style.display = '';
                document.getElementById("signin").style.display = 'none';

            } else {
                document.getElementById("imgleftmenu").style.display = 'none';
                document.getElementById("signin").style.display = '';
            }

            //******Here checking the IE version**********
            var rv = "others";
            if (navigator.appName == "Microsoft Internet Explorer") {
                var ua = navigator.userAgent;
                var re = new RegExp("MSIE ([0-9]{1,}[\.0-9]{0,})");
                if (re.exec(ua) != null)
                    rv = parseFloat(RegExp.$1);
            }

            rv = rv + " <%=MessagetoSave %>";
            //Calling the function,this save brower info,and other details
            getBrowserinfo(rv);


            if (isNaN(document.getElementById("spnTotalPrice").innerHTML)) {
            }
            else {
                //alert($("#spnTotalPrice").html() + "  " + $("#hdnRsCustomeServiceDetailsPrice").val());
                document.getElementById("spnTotalPrice").innerHTML = parseInt($("#spnTotalPrice").html()) + parseInt($("#hdnRsCustomeServiceDetailsPrice").val())
                //alert($("#spnTotalPrice").html() + "  " + $("#hdnRsCustomeServiceDetailsPrice").val());
            }

        </script>

        <!--//*******************Code to save data in webservice 20-08-2012*****************************************-->
        <script type="text/javascript" src="js/extrenalcallfromJson.js"></script>
        <!--*********End of Code to save data in webservice 20-08-2012**************************************************************-->

    </div>

    <%' This below link we need to call from here to get Airport and airline name %>
    <!--#include file="autocompleteByXML/XMLAutocompleteAirportAirline.asp"-->
    <div style="width: 900px;">
        <!--#include file="include_btm_footer.asp" -->
    </div>
</body>
</html>
