<!DOCTYPE html>
<html lang="en">
<head>
    <!--meta charset="utf-8"  //  This line causes the mailer to send two emails !! no idea !!-->
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <!--METADATA NAME="Microsoft ActiveX Data Objects 2.5 Library" TYPE="TypeLib" UUID="{00000205-0000-0010-8000-00AA006D2EA4}"-->
    <title>Data Entry Page</title>
    <meta name="description" content="Data Entry Page">
    <meta name="author" content="Rahul Mathur 10-29-15 Modified 7-11-18 Rahul Mathur">
    <!-- Style Sheets -->
    <!-- Don't remove or edit Jquerycclr.css. It contains boiler plate, normalizer and jquery styles-->
    <link href='../css/Jquerycclr.css' rel='stylesheet' type='text/css' />
    <!-- cclr Site CSS -->
    <link href='../css/cclrSiteStyles.css' rel='stylesheet' type='text/css' />
    <!-- Don't remove or edit DataEntryStyles.css. Add in style sheet for fabtabulous/DataEntry-->
    <link href='../css/DataEntryStyles.css' rel='stylesheet' type='text/css' />
    <!-- Add in script tag to make Prototype/fabtabulous Objects available Needs to be at top and bottom (EmailFoot.js) so it plays nice with jQuery???? -->
    <script type="text/javascript" src="../js/ProtoDataEntry.js"></script>
    <%@  language="JScript" %>
    <!--#include file="../ConnectionString/Connection.asp" -->
    <script language='JavaScript' runat='server'>
        // Initialize variables
        var NavSwitch = Request('NavSwitch');
        var Program = Request('Program');
        var LabRec = null;
        var ProgState = null;
        var Now = new Date();
        var TimeEntered = new Date() + '';
        TimeEntered = TimeEntered.replace(/:[0-9][0-9] EDT/, '');
        var ErrorCount = 0;
        var Error = new Array();
        var ErrorTest = new Array();
        var TabCount = 10;
        //This is used for the data entry accessed from consel.asp used by Admin staff.
        if (Session('Labnum') == 'cclr' && Session('Pinnum') == Session('Labnum')) {
            var Labnum = Request('Labnum');
        }
        else  // This nested if redirects to index page if session variables and cookies are expired.
        {
            if (Request.Cookies('Labnum') == null) {
                Response.Redirect('../default.html');
            }
            Session('Labnum') = Request.Cookies('Labnum');
            var Labnum = Request.Cookies('Labnum');
        }

        // Set Default Validation Email and Remarks validation for client and server sides.
        var DefaultRegExpress = '^[^`~#%_=+{}|:;</>@/$/^/*]{0,200}$';
        var DefaultErrorMess = 'Please do not exceed 200 characters or use any of these `~#%_=+{}|:;<>/$^*@';
        var EmailRegExpress = '^[^`~#%=+{}|:;</>?/$/^/*]*$';
        var EmailErrorMess = 'Pleas do not use these characters `~#%=+{}|:;<>?/$^*';
        var RemarksRegExpress = '^[^`~#%_=+{}|:;</>@/$/^/*]{0,1000}$';
        var RemarksErrorMess = 'Pleas do not exceed 1000 characters or use any of these  `~#%_=+{}|:;<>/$^*@';
        //************************************************************************************************
        // This include file has all the functions and the code that calls them.
        // It must be located here.
    </script>
    <!--#include file="../js/DataEntryFunctions.js"-->
    <script language='JavaScript' runat='server'>
        //************************************************************************************************
        // Start of code that sets up the heading for the error & conformation Emails for all programs.
        // Begining of Ebody the body of the email
        // Ebody is a variable that holds all the HTML to be Emailed it builds on itself line by line.
        var EHeading = "<table class='table table-bordered table-striped'><tr><td colspan='5'>Data submitted for " + Material + " samples " + "<b>" + LabRec('OddSample') + "</b> & <b>" + EvenSample + "</b> on <b> " + Now + "</font>" + "</b>.</td></tr>";
        if (NavSwitch == 'EnterData') {
            EHeading += "<tr><td colspan='5' align='center'><h1><font size='6' color='#800000'>*****Lab navigated to Page at " + Now + " !!!</font></h1></td></tr>";
            Subject = '*****Nav Alert**** For';
        }
        else if (NavSwitch == 'DataNotEntered') {
            EHeading += "<tr><td colspan='5'><h1><font size='6' color='#800000'>*****WARNING YOUR DATA SUBMISSION IS INCOMPLETE***** Please resubmit your data correcting any BadData fields listed below!!!</font></h1></td></tr>";
            Subject = '*****WARNING ***** Incomplete Data Submitted by';
        }
        else {
            var Subject = 'Confirmation of Data Submitted by ';
        }
        EHeading += "<tr><td colspan='5'align='left'>Data submitted by <b><font color='#000080'>" + LabRec('LN') + "</font> Lab # <font color='#000080'>" + LabRec('LabNum') + ".</font></b></td></tr>";
        EHeading += "<tr><td colspan='5'>Located in: <b><font color='#000080'>" + LabRec('City') + ", " + LabRec('State') + "</font></b></td></tr>";
        EHeading += "<tr><td colspan='5' align='center'><font size='4' color='#800000'>TEST DATA</font></td></tr>";
        EHeading += "<tr><td align='center'>Test Name</td><td align='center'>Test Unit</td><td align='center'>Sample No. " + LabRec('OddSample') + "</td><td align='center'>Sample No. " + EvenSample + "</td><td align='center'>Test Number</td></tr>";

        Ebody = EHeading + Ebody

        Ebody += "<tr><td colspan='5' align='left'>Remarks: <font color='#000080'>" + LabRec('Remarks') + "</font></td></tr>";
        Ebody += "<tr><td colspan='5' align='left' font size='1'><sub>Please do not reply to this email confirmation.  You may contact us at <a href=''></a>.</sub></font></td></tr>";
        Ebody += "</table>"
        // End of Ebody the body of the email

        // Beging of code thats creates the mailer object and sends the emails.
        var Mailer = Server.CreateObject('SMTPsvg.Mailer');

        // If statment to get around that some programs Sample Type (TestTitle(2) is null.
        //if(Program == 'Concrete' || Program == 'Rebar' || Program == 'MasMort' || Program == 'MasCem')
        if (SampleType == '') {
            Mailer.FromName = 'cclr PSP ' + Material + ' Data Confirmation';
            Mailer.Subject = '' + Subject + ' Laboratory # ' + LabRec('LabNum') + ' for ' + Material + ' Samples' + ' ' + LabRec('OddSample') + ' & ' + EvenSample + '.';
        }
        else {
            Mailer.FromName = 'cclr PSP ' + Material + ' ' + SampleType + ' Data Confirmation';
            Mailer.Subject = '' + Subject + ' Laboratory # ' + LabRec('LabNum') + "'s " + Material + ' ' + SampleType + ' Samples' + ' ' + LabRec('OddSample') + ' & ' + EvenSample + '.';
        }
        Mailer.FromAddress = ''

        //*************************************************************************************************
        //Use this mailer when on the verio site.
        //Mailer.RemoteHost = 'mail-fwd'

        //Use this mailer when on the applled inovations site.
        Mailer.RemoteHost = '';
        //*************************************************************************************************
        Mailer.ContentType = 'text/html';
        // If statement that determines if mail recipient gets emailed.
        if (NavSwitch == 'DataEntered' || NavSwitch == 'DataNotEntered') {
            Mailer.AddRecipient('', LabRec('Email'));
            //Mailer.Addcc ('',''+ Program +'@cclr.us, + LabRec('SecEmail'))  
            //Mailer.Addcc ('',''+ LabRec('SecEmail'))
            Mailer.Addcc('', '' + LabRec('SecEmail') + ',' + Program + '@cclr.us');
        }
        if (NavSwitch == 'EnterData') {
            Mailer.Addcc('', '' + Program + '@cclr.us,');
        }
        Mailer.BodyText = Ebody;
        if (NavSwitch != 'ViewData' && LabRec('LabNum') != 0) {
            Mailer.sendmail();
        }
        // End of code that creates the mailer object and sends the emails.
        // End of code that sends the error & conformation Emails for all programs.

        //Begin of code that generates the HTML heading 'DHeading'.
        //Heading       
        var DHeading = "<title>" + Material + " " + SampleType + "Data Entry Page</title>";
        DHeading += "<div id = 'NoPrint'><div id='includedHeader'></div></div>";
        DHeading += "<div class='container'><div id = 'NoPrint'><div class='float-right'>";
        DHeading += "<a href='ProgSelect.asp?NavSwitch=ProgSelect'><button type='submit' class='btn btn-primary'>Program Selection Page</button></a>";
        DHeading += "</div></div>";
        DHeading += "<table border='0' cellpadding='0' cellspacing='0' width='100%'>";
        DHeading += "<tr><td class='Right'>Laboratory Name:</td><td>&nbsp;</td><td class='DataLeft'>";
        DHeading += LabRec('LN') + "</td></tr>";
        DHeading += "<tr><td class='Right'>City:</td><td>&nbsp;</td><td class='DataLeft'>";
        DHeading += LabRec('City') + "</td></tr>";
        if (LabRec('State') != '') {
            DHeading += "<tr><td class='Right'>State:</td><td>&nbsp;</td><td class='DataLeft'>";
            DHeading += LabRec('State') + "</td></tr>";
        }
        var Country = LabRec('Country') + '';
        if (Country != '' && Country != 'null') {
            DHeading += "<tr><td class='Right'>Country:</td><td>&nbsp;</td><td class='DataLeft'>";
            DHeading += LabRec('Country') + "</td></tr>";
        }
        DHeading += "<tr><td class='Right'>Laboratory Number:</td><td>&nbsp;</td><td class='DataLeft'>";
        DHeading += LabRec('Labnum') + "</td></tr>";
        DHeading += "</table>";
        // Begin main form and table
        DHeading += "<form id='DataEntry' method='POST'onSubmit='return emailCheck(this.Email.value)' action='DataEntry.asp?NavSwitch=DataNotEntered&Program=" + Program + "&Labnum=" + Labnum + "')>";

        DHeading += "<table class='table table-bordered table-striped'>";
        if (Session('Labnum') == 'cclr' && Session('Pinnum') == '4705') {
            DHeading += "<tr><td colspan='5' class='alert'>This data was entered by cclr staff!</td>";
        }
        else if (Now < OpenDate) {
            DHeading += "<tr><td colspan='5' class='Title'>The Opening Date For " + Program + " PSP Sample Nos. " + LabRec('OddSample') + " and " + EvenSample + " is  <u>" + OpenDateStr + "</u> !</td></tr>";
            DHeading += "<tr><td colspan='5'class='Title'>The data below is laboratory " + LabRec('Labnum') + "'s data for Samples Nos. " + LabRec('OddSample') + " and " + EvenSample + ".</td></tr>";
        }

        else if (Now >= OpenDate && Now < CloseDate) {
            if (NavSwitch == 'EnterData') {

                if (SampleType == '') {
                    DHeading += "<tr><td colspan='5' class='Title'>Please Submit Your Data for " + Material + " PSP Sample Nos. " + LabRec('OddSample') + " and " + EvenSample + "</td></tr>";
                }
                else {
                    DHeading += "<tr><td colspan='5' class='Title'>Please Submit Your Data for " + Material + " " + SampleType + " PSP Sample Nos. " + LabRec('OddSample') + " and " + EvenSample + "</td></tr>";
                }
            }
            else if (NavSwitch == 'DataNotEntered') {
                DHeading += "<tr><td colspan='5' class='Alert'><u>ALERT !!  YOUR DATA HAS NOT BEEN SUBMITTED !!</u></td></tr>";
                DHeading += "<tr><td colspan='5' class='Alert'>PLEASE SCROLL DOWN AND CORRECT THE " + ErrorCount + " ERROR(S) ON THE PAGE AND RESUBMIT DATA!!</td></tr>";
                for (j = 1; j < ErrorCount + 1; j++) {
                    DHeading += "<tr><td colspan='5' class='alert'><font color='#CE2648'>Error # " + j + " on " + ErrorTest[j] + " " + Error[j] + "</td></tr>";
                }
            }
            else if (NavSwitch == 'DataEntered') {
                if (LabRec('SecEmail') == " ") {
                    DHeading += "<tr><td colspan='5' class='Title'><h4>A confirmation Email has been sent to the Email address <span class='Data'> " + LabRec('Email') + "</span> !</td></tr>";
                }
                else {
                    DHeading += "<tr><td colspan='5' class='Title'><h4>A confirmation Email has been sent to the Email address <span class='Data'>" + LabRec('Email') + " & " + LabRec('SecEmail') + "</span> !</td></tr>";
                }
                DHeading += "<tr><td colspan='5'class='Title'>You should retain this Email for your records!</h4></td>";
            }
        }
        //if (Now.setTime(Now.getTime()) > CloseDate)
        else if (Now > CloseDate)
        //if (Now.setTime(Now.getTime()) > CloseDate)
        {
            DHeading += "<tr><td colspan='5' class='Title'>The closing date for PSP Samples Nos. " + LabRec('OddSample') + " and " + EvenSample + " was 12:00 am, <u>" + CloseDateStr + "</u> !</td></tr>";
            DHeading += "<tr><td colspan='5' class='Title'>The data below is the current data for laboratory number <span class='Data'>" + LabRec('Labnum') + "</span>.</td></tr>";
        }

        // Begin of code that generates the HTML code for Email.
        if (NavSwitch == 'EnterData' || NavSwitch == 'DataNotEntered') {
            //Response.Write("<tr><td colspan='3' width='100%' height='20' align='center' ><input type='submit' value='Submit Data' name='SubmitT1' tabindex='1'></td></tr>";
            DHeading += "<tr><td colspan='5' class='Center'><button type='submit' class='btn btn-primary 'tabindex='1'>Submit Data</button></td></tr>";
            DHeading += "<tr><td>";
            DHeading += "Email Address:";
            DHeading += "</td><td colspan='4'>";
            DHeading += "<input type='text' name='Email' class='Email' title='Enter Email addresss' tabindex='8' value='" + fV.elements['Email'].elementValue.split(' ').join('') + "'>";
            DHeading += "</td></tr><tr><td>";
            DHeading += "Second Email Address: <span class='small'>(Optional)</span>";
            DHeading += "</td><td colspan='4'>";
            DHeading += "<input type='text' name='SecEmail' class='SecEmail' title='Enter Second Email addresss (opt)' tabindex='9' value='" + fV.elements['SecEmail'].elementValue.split(' ').join('') + "'>";
            DHeading += "</td></tr>";
        }
        else {
            DHeading += "<tr><td>";
            DHeading += "Email Address:";
            DHeading += "</td><td colspan='4' class='Email'>";
            DHeading += LabRec('Email');
            DHeading += "</td></tr><tr><td>";
            DHeading += "Second Email Address:";
            DHeading += "</td><td colspan='4' class='Email'>";
            DHeading += LabRec('SecEmail');
            DHeading += "</td></tr>";
        }
        //End of code that generates the HTML code for Email.     
        DHeading += "<tr><td colspan='5' class='Title'>Test Results: Report as indicated in ( )</td></tr>";
        // End of code that generates the HTML heading.

        // Begining of the HTML ending code used for all Programs.
        var DEnding = '';
        DEnding += "<tr><td colspan='5' class='Remark'><span class='Title'>Remarks:</span>";
        if (NavSwitch == 'EnterData' || NavSwitch == 'DataNotEntered') {
            var FieldLength = fV.elements['Remarks'].elementValue.length;
            if (NavSwitch == 'EnterData') { FieldLength-- };
            DEnding += "<textarea  name='Remarks' class='Remarks' title='Enter Your Remarks' tabindex='200'>" + fV.elements['Remarks'].elementValue + "</textarea>";
        } else {
            DEnding += LabRec('Remarks');
        }
        DEnding += "</td></tr>";
        if (NavSwitch == 'EnterData' || NavSwitch == 'DataNotEntered') {
            DEnding += "<tr><td colspan='5' align='center' ><button type='submit' class='btn btn-primary 'tabindex='210'>Submit Data</button></td></tr>";
        }
        DEnding += "</table></form></div><div id='includedFooter'></div></body></html>";
        // End of server side code.
        // End of the HTML ending code used for all Programs.

        //close the record sets and release them.
        LabRec.close();
        LabRec = null;
        Close_Connection();
        // Write the heading body and ending to the screen!
        Response.Write(DHeading);
        // DBody is created in the DataEntry Functions code.
        Response.Write(DBody);
        Response.Write(DEnding);
        // unlock the application
        Application.Unlock();
    </script>
    <!-- Don't remove or edit JquerySite.js. It contains jquery library dataTable and many other jquery extentions-->
    <!-- Placed at the end of the document so the pages load faster -->
    <script type="text/javascript" language="javascript" src="../js/JquerySite.js"></script>
    <!-- jquery scripts to include header and footer files-->
    <script>jQuery(function(){jQuery("#includedHeader").load("../Header.html");});</script>
    <script>jQuery(function(){jQuery("#includedFooter").load("../Footer.html");});</script>
    <!--Add in script tag to make Prototype/fabtabulous Objects available Needs to be at top and bottom so it plays nice with jQuery -->
    <script type="text/javascript" src="../js/ProtoDataEntry.js"></script>
    <!--  End Bootstrap core JavaScript-->
    <script type="text/javascript">
        // Begining of client side code that sets Up objects for client form validation.
        function formCallback(result, form) { window.status = "validation callback for form '" + form.id + "': result = " + result; }
        var valid = new Validation('DataEntry', { immediate: true, onSubmit: false, onFormValidate: formCallback });
        //Next four lines creates the validator object for Email, SecEmail, Remarks and Default for client side using a little ASP.
        Validation.add('Default', '<%=DefaultErrorMess%>', { pattern: new RegExp('<%=DefaultRegExpress%>') });
        Validation.add('Email', '<%=EmailErrorMess%>', { pattern: new RegExp('<%=EmailRegExpress%>') });
        Validation.add('SecEmail', '<%=EmailErrorMess%>', { pattern: new RegExp('<%=EmailRegExpress%>') });
        Validation.add('Remarks', '<%=RemarksErrorMess%>', { pattern: new RegExp('<%=RemarksRegExpress%>') });
        var TestJson = JSON.parse('<%=TestJson%>'); // this parse does not work in IE 8.0.6 or older
        // This for loops through Test object and creates the client validation objects for special fields,
        for (var j = 0; j <= TestJson.Count[0].ClientSideCount; j++) {
            Validation.add('Test' + TestJson.Test[j].TestNum, TestJson.Test[j].ErrorMess, { pattern: new RegExp(TestJson.Test[j].RegExpress) });
        }
        // End of code that sets Up objects for client form validation.
        //This line calls the client side validation after form has been submitted.
        if ('' + <%= NavSwitch %> + '' != 'EnterData');
        {
            var result = valid.validate()
        }
        // Begin client side email validator function
        function emailCheck(emailStr) {
            /* The following variable tells the rest of the function whether or not
            to verify that the address ends in a two-letter country or well-known
            TLD.  1 means check it, 0 means don't. */
            var checkTLD = 1;
            /* The following is the list of known TLDs that an e-mail address must end with. */
            var knownDomsPat = /^(com|net|org|edu|int|mil|gov|arpa|biz|aero|name|coop|info|pro|museum)$/;
            /* The following pattern is used to check if the entered e-mail address
            fits the user@domain format.  It also is used to separate the username
            from the domain. */
            var emailPat = /^(.+)@(.+)$/;
            /* The following string represents the pattern for matching all special
            characters.  We don't want to allow special characters in the address.
            These characters include ( ) < > @ , ; : \ " . [ ] */
            var specialChars = "\\(\\)><@,;:\\\\\\\"\\.\\[\\]";
            /* The following string represents the range of characters allowed in a
            username or domainname.  It really states which chars aren't allowed.*/
            var validChars = "\[^\\s" + specialChars + "\]";
            /* The following pattern applies if the "user" is a quoted string (in
            which case, there are no rules about which characters are allowed
            and which aren't; anything goes).  E.g. "jiminy cricket"@disney.com
            is a legal e-mail address. */
            var quotedUser = "(\"[^\"]*\")";
            /* The following pattern applies for domains that are IP addresses,
            rather than symbolic names.  E.g. joe@[123.124.233.4] is a legal
            e-mail address. NOTE: The square brackets are required. */
            var ipDomainPat = /^\[(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})\]$/;
            /* The following string represents an atom (basically a series of non-special characters.) */
            var atom = validChars + '+';
            /* The following string represents one word in the typical username.
            For example, in john.doe@somewhere.com, john and doe are words.
            Basically, a word is either an atom or quoted string. */
            var word = "(" + atom + "|" + quotedUser + ")";
            // The following pattern describes the structure of the user
            var userPat = new RegExp("^" + word + "(\\." + word + ")*$");
            /* The following pattern describes the structure of a normal symbolic
            domain, as opposed to ipDomainPat, shown above. */
            var domainPat = new RegExp("^" + atom + "(\\." + atom + ")*$");
            /* Finally, let's start trying to figure out if the supplied address is valid. */
            /* Begin with the coarse pattern to simply break up user@domain into
            different pieces that are easy to analyze. */
            var matchArray = emailStr.match(emailPat);
            if (matchArray == null) {
                /* Too many/few @'s or something; basically, this address doesn't
                even fit the general mould of a valid e-mail address. */
                alert("Your Email address can not be blank, or it seems incorrect (check @ and .'s)");
                return false;
            }
            var user = matchArray[1];
            var domain = matchArray[2];
            // Start by checking that only basic ASCII characters are in the strings (0-127).
            for (i = 0; i < user.length; i++) {
                if (user.charCodeAt(i) > 127) {
                    alert("Ths username contains invalid characters.");
                    return false;
                }
            }
            for (i = 0; i < domain.length; i++) {
                if (domain.charCodeAt(i) > 127) {
                    alert("Ths domain name contains invalid characters.");
                    return false;
                }
            }

            // See if "user" is valid
            if (user.match(userPat) == null) {
                // user is not valid
                alert("The username doesn't seem to be valid.");
                return false;
            }

            /* if the e-mail address is at an IP address (as opposed to a symbolic
            host name) make sure the IP address is valid. */
            var IPArray = domain.match(ipDomainPat);
            if (IPArray != null) {
                // this is an IP address
                for (var i = 1; i <= 4; i++) {
                    if (IPArray[i] > 255) {
                        alert("Destination IP address is invalid!");
                        return false;
                    }
                }
                return true;
            }

            // Domain is symbolic name.  Check if it's valid.
            var atomPat = new RegExp("^" + atom + "$");
            var domArr = domain.split(".");
            var len = domArr.length;
            for (i = 0; i < len; i++) {
                if (domArr[i].search(atomPat) == -1) {
                    alert("The domain name does not seem to be valid.");
                    return false;
                }
            }

            /* domain name seems valid, but now make sure that it ends in a
            known top-level domain (like com, edu, gov) or a two-letter word,
            representing country (uk, nl), and that there's a hostname preceding
            the domain or country. */
            if (checkTLD && domArr[domArr.length - 1].length != 2 &&
                domArr[domArr.length - 1].search(knownDomsPat) == -1) {
                alert("The address must end in a well-known domain or two letter " + "country.");
                return false;
            }

            // Make sure there's a host name preceding the domain.

            if (len < 2) {
                alert("This address is missing a hostname!");
                return false;
            }

            // If we've gotten this far, everything's valid!
            return true;
        }
// End client side email validator function///
// End of all client side code.
    </script>
