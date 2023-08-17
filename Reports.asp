<%@  language="JScript" %>
<!--#include file="../ConnectionString/Connection.asp" -->
<script language="JavaScript" runat="server">	
    //<created 11-2-09 by Rahul Mathur: Last Modified: 1-6-17 by Rahul Mathur></created>
    // Declare variables 
    var Program = Request('Program');
    var NavSwitch = Request('NavSwitch');
    // Next line is there to make this old page work with new system
    if (NavSwitch == 'Ratings') { var NavSwitch = 0 };
    var TestNumber = Request('TestNumber');
    var UnitSwitch = Request('UnitSwitch');
    var RoundSwitch = Request('RoundSwitch');
    // This keeps smarties from being able to view the data of others.
    if (Session('Pinnum') == null && Request('ChoseLab') != -1) {
        var ChoseLab = Session('Labnum');
    }
    else {
        var ChoseLab = Request('ChoseLab');
    }
    var ChoseSample = Request('ChoseSample');
    var EvenSample = ChoseSample;
    EvenSample = EvenSample++;
    var NumLabs = 0;
    var OddSum = 0;
    var EvenSum = 0;
    var SampCount = 0;
    var SampListCount = 0;
    var Unit = null;
    var Converter = null;
    var LabDataOdd = null;
    var LabDataEven = null;
    var OddAvg = null;
    var EvenAvg = null;
    var OddSD = null;
    var EvenSD = null;

    // Declare arrays
    var OddSample = [];
    var OddSamples = [];
    var EvenSample = [];
    var SampleAvg = [];
    var OddData = [];
    var EvenData = [];
    var OddSort = [];
    var EvenSort = [];
    var OddSD = [];
    var EvenSD = [];
    var OddAvg = [];
    var EvenAvg = [];
    var OddZ = [];
    var EvenZ = [];
    var LabRec = [];
    var MultiLabNum = [];
    var ProgName = [];
    var ProgInfo = [];
    var WeirdData = [];
    var TestStats = [];
    var SampleStats = [];
    var XTicks = [];
    var Labs = [];
    var LabData = [];
    var SampNos = [];
    var SampList = [];

    // Functions
    // Begining of Stats function
    // Call this function to do all the stats ( Avg SD CV )  
    // Input an array (Data); out puts object Result with three properties Result.AVG, Result.SD, Result.CV		
    function Stats(Data) {
        var Result = { Avg: 0, SD: 0, CV: 0 }, t = Data.length;
        for (var m, s = 0, l = t; l--; s += parseFloat(Data[l]));
        for (m = Result.Avg = s / t, l = t, s = 0; l--; s += Math.pow(Data[l] - m, 2));
        Result.SD = Math.sqrt(s / t)
        return Result.CV = 100 * Result.SD / Result.Avg, Result
    }
    // End of Stats function

    // Begining of DecimalPlaces function
    function DecimalPlaces(Number) // Returns the number of decimal places. Used with RoundASTM function
    {
        var Num = new String(); Num = '' + Number;
        var Pos = 0; var Decimals = 0;
        while (Num.substring(Pos - 1, Pos) !== ".") { if (Pos == Num.length) { return Decimals }; Pos += 1; }
        while (Pos < (Num.length)) { Decimals++; Pos += 1; } return Decimals;
    }
    // End of DecimalPlaces function

    // Begining of function RoundASTM()-->Rounds to one more decimal place than decimal places displayed.
    function RoundASTM(Number, Decimals) {
        result = Number
        var i = parseFloat(Number);
        if (isNaN(i)) { i = 0.00; }
        var minus = '';
        if (i < 0) { minus = '-'; }
        i = Math.abs(i);
        if (Decimals == -1) { i = parseInt((i + 5) * 0.1); i = i / 0.1; }
        else if (Decimals == 0) { i = parseInt(i + .5); }
        else if (Decimals == 1) { i = parseInt((i + .05) * 10); i = i / 10; }
        else if (Decimals == 2) { i = parseInt((i + .005) * 100); i = i / 100; }
        else if (Decimals == 3) { i = parseInt((i + .0005) * 1000); i = i / 1000; }
        else if (Decimals == 4) { i = parseInt((i + .00005) * 10000); i = i / 10000; }
        else if (Decimals == 5) { i = parseInt((i + .000005) * 100000); i = i / 100000; }
        else if (Decimals == 6) { i = parseInt((i + .0000005) * 1000000); i = i / 1000000; }
        RoundedNum = new String(i);
        RoundedNum = minus + RoundedNum;
        return RoundedNum;
    }
    // End of function RoundASTM()
    // End functions	

    if (TheConn == null) {
        Response.Write("<p> Error: Database Connection Error");
    }
    else {
        CmdTextStr = 'DECLARE @Program AS Char(10);'
        CmdTextStr += 'SELECT @Program = ?;'
        CmdTextStr += 'SELECT Program, Material, SampleType, CloseDate, TestNumber, TestTitle, DisplayOrder, DisplayStatus, TestUnit, TestDec FROM ProgramInfo Where (Program = @Program);'
        TheComm = Server.CreateObject("ADODB.Command");// Create a command object.WHERE (OddSample = @SampleNumber) 
        TheComm.ActiveConnection = TheConn;// Assine the comand object to the connectin object TheConn.
        TheComm.CommandText = CmdTextStr;// Assine the the comand text to the command object TheComm.	
        // Use parameterized query to stop SQL Injection attacks
        TheComm.Parameters.Append(TheComm.CreateParameter('@Program', 129, 1, 10, Program)); // Create the parameter query based on the variable Program. Can be Concrete, Portland Cement ect. 
        RecordSet = TheComm.Execute(adCmdText + adExecuteNoRecords);// Record set from parameterized query table "ProgramInfo" to retrieve heading info for the selected program.
        var ProgTest = [];
        //This while reads in the data from a specific program in ProgInfo table of pspdata data base into an array of arrays "ProgTest".
        g = 0
        while (!RecordSet.EOF) {
            ProgTest[g] = [];
            ProgTest[g]['Program'] = RecordSet.Fields('Program').value;
            ProgTest[g]['Material'] = RecordSet.Fields('Material').value;
            ProgTest[g]['SampleType'] = RecordSet.Fields('SampleType').value;
            ProgTest[g]['CloseDate'] = RecordSet.Fields('CloseDate').value;
            ProgTest[g]['TestNumber'] = RecordSet.Fields('TestNumber').value;
            ProgTest[g]['DisplayOrder'] = RecordSet.Fields('DisplayOrder').value;
            ProgTest[g]['DisplayStatus'] = RecordSet.Fields('DisplayStatus').value;
            ProgTest[g]['TestTitle'] = RecordSet.Fields('TestTitle').value;
            ProgTest[g]['TestUnit'] = RecordSet.Fields('TestUnit').value;
            ProgTest[g]['TestDec'] = RecordSet.Fields('TestDec').value;
            RecordSet.MoveNext;
            g++
        }
        RecordSet.close();
        ErrorProg = ProgTest[0]['Material'] + " " + ProgTest[0]['SampleType'];
        NumOfTests = g - 1
        Material = ProgTest[0]['Material']
        SampleType = ProgTest[0]['SampleType']
    }
    //********************// html block for all reports but not charts**************   
    Response.Write("<html><head><title>" + Material + " " + SampleType + " Report Page</title>");
    // Style Sheets 
    // Boiler plate normalizer css
    Response.Write("<link href='../css/Jquerycclr.css' rel='stylesheet' type='text/css'/>");
    // cclr Site CSS
    Response.Write("<link href='../css/cclrSiteStyles.css' rel='stylesheet' />");
    Response.Write("<style type = 'text/css' media = 'print'>#NoPrint{ display: none;}</style>");
    Response.Write("<div id='includedHeader'></div>");
    Response.Write("</div>");
    Response.Write("<div class='container'>");
    //************************************************************************************************************************
    if (NavSwitch < 6)// Dispays Ratings page and Muti Ratings page.
    {
        if (Session('Labnum') == null) {
            if (Request.Cookies('Labnum') == null) {
                Response.Redirect('../default.asp')
            }
            else {
                Session("Labnum") = Request.Cookies('Labnum') + '';
            }
        }
        // The function that desplays all the ratings
        function fRatings(Program, Labnum) {
            // Define varables and arrays
            var ok = false;
            var ChoseLab;
            var LabNo = [];
            var SampleNo = [];
            var EvenSampleNo = [];
            var LabField = ""
            var RightRec = new Array()

            // Get the stats for a sample pair
            var TheComm = Server.CreateObject("ADODB.Command")
            TheComm.ActiveConnection = TheConn
            //var LatestSample = 177
            var Pinnum = Session('Pinnum');
            Pinnum = Pinnum + '';
            if (Pinnum.length == 5) {
                if (Pinnum == 88888) {
                    var SelectRec = "SELECT * FROM " + Program + "Rate;"
                    TheComm.CommandText = SelectRec;
                }
                else {
                    var TheComm2 = Server.CreateObject("ADODB.Command")
                    TheComm2.ActiveConnection = TheConn
                    var SelectMultiLabs = "DECLARE @ReceiverPin AS Int;"
                    SelectMultiLabs += "SELECT @ReceiverPin = ?;"
                    SelectMultiLabs += "SELECT * FROM MultiLabs WHERE (ReceiverPin = @ReceiverPin);"
                    TheComm2.CommandText = SelectMultiLabs;
                    TheComm2.Parameters.Append(TheComm2.CreateParameter('@ReceiverPin', 3, 1, 5, Session('Pinnum')));
                    var MultiLabNum = TheComm2.Execute(adCmdText + adExecuteNoRecords);
                    TheComm2 = null;

                    var DeclareRec = "DECLARE @Labnum AS Char(5)"
                    var ParamRec = "SELECT @Labnum = ?"
                    var SqlRec = "SELECT * FROM " + Program + "Rate WHERE (Labnum = @Labnum)"
                    TheComm.Parameters.Append(TheComm.CreateParameter('@Labnum', 129, 1, 4, MultiLabNum('SenderLabno')));
                    MultiLabNum.MoveNext
                    var Count = 1
                    while (!MultiLabNum.EOF) {
                        DeclareRec += ",@Labnum" + Count + " AS Char(5)"
                        ParamRec += ",@Labnum" + Count + " = ?"
                        SqlRec += " OR (Labnum = @Labnum" + Count + ")"
                        TheComm.Parameters.Append(TheComm.CreateParameter('@Labnum' + Count, 129, 1, 5, MultiLabNum('SenderLabno')));
                        MultiLabNum.MoveNext
                        Count++
                    }
                    DeclareRec += ";"
                    ParamRec += ";"
                    SqlRec += ";"
                    var SelectRec = DeclareRec + ParamRec + SqlRec

                    TheComm.CommandText = SelectRec;
                }
            }
            else {
                // Using parameterized query
                var SelectRec = "DECLARE @Labnum AS Char(5);"
                SelectRec += "SELECT @Labnum = ?;"
                SelectRec += "SELECT * FROM " + Program + "Rate WHERE (Labnum = @Labnum);"
                TheComm.CommandText = SelectRec;
                TheComm.Parameters.Append(TheComm.CreateParameter('@Labnum', 129, 1, 5, Labnum));
            }
            var LabRec = TheComm.Execute(adCmdText + adExecuteNoRecords);
            FieldCount = LabRec.Fields.Count;
            if (LabRec.EOF) {
                ok = false;
                return ok;
            }
            else {
                ok = true;
            }

            // Start of crazy logic that rebuilds each select box based on the selection of the other
            if (NavSwitch == 0) {
                var LatestSampleNo = [];
                //LabRec.MoveFirst;
                var LabRec = TheComm.Execute(adCmdText + adExecuteNoRecords);
                LatestSampleCount = 0;
                while (!LabRec.EOF) {
                    LatestSampleNo[LatestSampleCount] = LabRec.Fields('OddSample').Value;
                    LabRec.MoveNext;
                    LatestSampleCount = LatestSampleCount + 1
                }
                LatestSampleNo.sort(function (a, b) { return b - a; });
                ChoseSample = LatestSampleNo[0]
            }
            else {
                ChoseSample = Request('ChoseSample')
            }
            var LabRec = TheComm.Execute(adCmdText + adExecuteNoRecords);
            LabCount = 0;
            while (!LabRec.EOF) {
                if (LabRec.Fields('OddSample').Value == ChoseSample) {
                    LabNo[LabCount] = LabRec.Fields('Labnum').Value;
                    LabRec.MoveNext;
                    LabCount = LabCount + 1
                }
                else {
                    LabRec.MoveNext;
                }
            }
            LabNo.sort(function (a, b) { return a - b; });
            // remove duplicates
            var Temp = [];
            for (var x = 0; x < LabNo.length; x++) {
                var isDup = false;
                for (var y = 0; y < Temp.length; y++) {
                    if (Temp[y] == LabNo[x]) {
                        isDup = true;
                        break;
                    }
                }
                if (!isDup) Temp[Temp.length] = LabNo[x];
            }
            LabNo = Temp
            var LabRec = TheComm.Execute(adCmdText + adExecuteNoRecords);
            SampleCount = 0;
            while (!LabRec.EOF) {
                if (NavSwitch == 0) { ChoseLab = LabNo[0] }
                else { ChoseLab = Request('ChoseLab') }
                if (LabRec.Fields('Labnum').Value == ChoseLab) {
                    SampleNo[SampleCount] = LabRec.Fields('OddSample').Value;
                    LabRec.MoveNext;
                    SampleCount = SampleCount + 1
                }
                else {
                    LabRec.MoveNext;
                }
            }
            SampleNo.sort(function (a, b) { return b - a; });
            // remove duplicates
            var Temp = [];
            for (var x = 0; x < SampleNo.length; x++) {
                var isDup = false;
                for (var y = 0; y < Temp.length; y++) {
                    if (Temp[y] == SampleNo[x]) {
                        isDup = true;
                        break;
                    }
                }
                if (!isDup) { Temp[Temp.length] = SampleNo[x] };
            }
            SampleNo = null
            var SampleNo = [];
            for (var x = 0; x < Temp.length; x++) {
                SampleNo[x] = Temp[x]
            }
            Temp = null
            var LabRec = TheComm.Execute(adCmdText + adExecuteNoRecords);
            for (j = 0; j < LabNo.length; j++) {
                if (ChoseLab == LabNo[j]) {
                    if (NavSwitch == 1) {
                        j = j + 1
                        if (j > LabNo.length - 1) { ChoseLab = LabNo[0] } else { ChoseLab = LabNo[j] };
                    }
                    else if (NavSwitch == 2) {
                        j = j - 1
                        if (j < 0) { ChoseLab = LabNo[LabNo.length - 1] } else { ChoseLab = LabNo[j] };
                    };
                    break;
                }
            }
            for (p = 0; p < SampleNo.length; p++) {
                if (ChoseSample == SampleNo[p]) {
                    if (NavSwitch == 3) {
                        p = p + 1
                        if (p > SampleNo.length - 1) { p = 0 }
                        ChoseSample = SampleNo[p]
                    }
                    else if (NavSwitch == 4) {
                        p = p - 1
                        if (p < 0) { p = SampleNo.length - 1 }
                        ChoseSample = SampleNo[p]
                    };
                    break;
                }
            }
            if (NavSwitch == 0) { ChoseLab = LabNo[0]; ChoseSample = LatestSampleNo[0]; }
            if (NavSwitch == 5) { ChoseLab = Request('ChoseLab'); ChoseSample = Request('ChoseSample') }
            LabRec.MoveFirst;
            while (!LabRec.EOF) {
                LabField = LabRec.Fields('Labnum').Value;
                SampleField = LabRec.Fields('OddSample').Value;
                if (ChoseLab == LabField && ChoseSample == SampleField) {
                    var EvenSample = ++SampleField;
                    FieldIndex = 0;
                    while (FieldIndex < FieldCount) {
                        RightRec[FieldIndex] = LabRec.Fields(FieldIndex).Value;
                        FieldIndex = FieldIndex + 1;
                    }
                }
                LabRec.MoveNext;
            }

            // Used to determin the number of decimal places the average and SD should have.            
            var SelectDecRec = "SELECT * From " + Program + "Rate"
            var DecRec = TheConn.Execute(SelectDecRec);

            // Connect to stats tabel and retreave tests stats.
            var SelectStats = "DECLARE @ChoseSample AS Int;"
            SelectStats += "SELECT @ChoseSample = ?;"
            SelectStats += "SELECT * FROM " + Program + "Stats WHERE (OddSample = @ChoseSample);"
            var TheComm2 = Server.CreateObject("ADODB.Command")
            TheComm2.ActiveConnection = TheConn
            TheComm2.CommandText = SelectStats;
            TheComm2.Parameters.Append(TheComm2.CreateParameter('@ChoseSample', 3, 1, 3, ChoseSample));
            var SampleStats = TheComm2.Execute(adCmdText + adExecuteNoRecords);
            var LabRec = TheComm.Execute(adCmdText + adExecuteNoRecords);

            Response.Write("<div id = 'NoPrint'>");
            Response.Write("<div class='float-right'>");
            Response.Write("<a href='ProgSelect.asp?NavSwitch=10'><button type='submit' class='btn btn-primary'>Program Selection Page</button></a>");
            Response.Write("</div>");

            if (Session('Pinnum') != null) {
                Response.Write("<table border='0'align='center'>");
                Response.Write("<form name='RatingsForm' METHOD='POST' ACTION='Reports.asp?Program=" + Program + "&UnitSwitch=" + UnitSwitch + "&RoundSwitch=" + RoundSwitch + "&NavSwitch=" + 5 + "'>");

                Response.Write("<tr>");
                Response.Write("<td align='center'>");
                Response.Write("<a href='Reports.asp?Program=" + Program + "&ChoseSample=" + ChoseSample + "&ChoseLab=" + ChoseLab + "&UnitSwitch=" + UnitSwitch + "&RoundSwitch=" + RoundSwitch + "&NavSwitch=" + 2 + "'target='_self'>");
                Response.Write("<span class='icon-backward2'></span></a>");
                Response.Write("</td>");

                Response.Write("<td align='center'>");
                Response.Write("<a href='javascript:window.print()'target='_self'>Print Page</a>");
                Response.Write("<br>");
                Response.Write("<select language='javascript' onChange='this.form.submit();' size='1' name='ChoseLab'>");
                Response.Write("<option selected  value='" + ChoseLab + "'>" + ChoseLab + '' + "</option>");
                for (j = 0; j < LabNo.length; j++) Response.Write("<font size='3'><option  value='" + LabNo[j] + "'>" + LabNo[j] + '' + "</option></b>");
                Response.Write("</select>");
                Response.Write("<br>&nbsp;");
                Response.Write("</td>");

                Response.Write("<td align='center'>");
                Response.Write("<a href='Reports.asp?Program=" + Program + "&ChoseSample=" + ChoseSample + "&ChoseLab=" + ChoseLab + "&UnitSwitch=" + UnitSwitch + "&RoundSwitch=" + RoundSwitch + "&NavSwitch=" + 1 + "'target='_self'>");
                Response.Write("<span class='icon-forward3'></span></a>");
                Response.Write("</td>");

                Response.Write("<td align='center'>");
                Response.Write("<a href='Reports.asp?Program=" + Program + "&ChoseSample=" + ChoseSample + "&ChoseLab=" + ChoseLab + "&NavSwitch=" + 3 + "'target='_self'><span class='icon-backward2'></span></a>");
                Response.Write("</td>");

                Response.Write("<td align='center'>");
                Response.Write("<a href='../Psp/Reports/" + Material + " Report " + ChoseSample + ".pdf 'target='_self'>View Report</a>");
                Response.Write("<br>");
                Response.Write("<select language='javascript' onChange='this.form.submit();' size='1' name='ChoseSample'>");
                Response.Write("<option selected  value='" + ChoseSample + "'>" + ChoseSample + " & " + EvenSample + '' + "</option>");
                for (j = 0; j < SampleNo.length; j++) Response.Write("<font size='3'><option  value='" + SampleNo[j] + "'>" + SampleNo[j] + " & " + ++SampleNo[j] + '' + "</option></b>");
                Response.Write("</select>");
                Response.Write("<br>&nbsp;");
                Response.Write("</td>");

                Response.Write("<td align='center'>");
                Response.Write("<a href='Reports.asp?Program=" + Program + "&ChoseSample=" + ChoseSample + "&ChoseLab=" + ChoseLab + "&UnitSwitch=" + UnitSwitch + "&RoundSwitch=" + RoundSwitch + "&NavSwitch=" + 4 + "'target='_self'>");
                Response.Write("<span class='icon-forward3'></span></a>");
                Response.Write("</td>");
                Response.Write("</tr>");

                Response.Write("</form>");
                Response.Write("</table>");
            }
            else {
                Response.Write("<table border='0' align='center'>");

                Response.Write("<tr>");

                Response.Write("<td>");
                Response.Write("<a href='Reports.asp?Program=" + Program + "&ChoseSample=" + ChoseSample + "&ChoseLab=" + ChoseLab + "&UnitSwitch=" + UnitSwitch + "&RoundSwitch=" + RoundSwitch + "&NavSwitch=" + 3 + "'target='_self'>");
                Response.Write("<span class='icon-backward2'>");
                Response.Write("</span></a>");
                Response.Write("</td>");

                Response.Write("<td align='center'>");
                Response.Write("<br>");
                Response.Write("<a href='javascript:window.print()'target='_self'>Print Page</a>");
                Response.Write("<br>");
                Response.Write("<form name='RatingsForm' METHOD='POST' ACTION='Reports.asp?Program=" + Program + "&NavSwitch=" + 5 + "'>");
                Response.Write("<select language='javascript' onChange='this.form.submit();' size='1' name='ChoseSample'>");
                Response.Write("<option selected  value='" + ChoseSample + "'>" + ChoseSample + " & " + EvenSample + '' + "</option>");
                for (j = 0; j < SampleNo.length; j++) Response.Write("<option  value='" + SampleNo[j] + "'>" + SampleNo[j] + " & " + ++SampleNo[j] + '' + "</option>");
                Response.Write("</select><input TYPE='hidden'name='ChoseLab' value='" + ChoseLab + "'>");

                Response.Write("<br>");
                Response.Write("<a href='../Psp/Reports/" + ProgInfo.Fields(1).Value + " Report " + ChoseSample + ".pdf 'target='_self'>View Report</a>");
                Response.Write("</form>");
                Response.Write("</td>");

                Response.Write("<td>");
                Response.Write("<a href='Reports.asp?Program=" + Program + "&ChoseSample=" + ChoseSample + "&ChoseLab=" + ChoseLab + "&UnitSwitch=" + UnitSwitch + "&RoundSwitch=" + RoundSwitch + "&NavSwitch=" + 4 + "'target='_self'>");
                Response.Write("<span class='icon-forward3'>");
                Response.Write("</span></a>");
                Response.Write("</td>");
                Response.Write("</tr>");

                Response.Write("</table>");
            }
            Response.Write("</div>"); // No print				

            // Start of code that generates the html code
            // this desplays the labs ratings           
            Response.Write("<h3>" + "Report Date: " + SampleStats.Fields('FinalRptDate') + "</h3>");
            Response.Write("<h3>Results For Laboratory No.&nbsp;" + ChoseLab + "</h3>");

            if (Program == 'Concrete' || Program == 'Rebar' || Program == 'MasMort' || Program == 'MasCem') {
                Response.Write("<h3>" + "cclr " + Material + " Proficiency Sample Nos. " + ChoseSample + " & " + EvenSample + "</h3>");
            }
            else {
                Response.Write("<h3>" + "cclr " + Material + " " + SampleType + " Proficiency Sample Nos. " + ChoseSample + " & " + EvenSample + "</h3>");
            }
            Response.Write("<table class='table table-bordered table-striped'>");
            Response.Write("<caption>");
            Response.Write("<div id = 'NoPrint'>");
            if (RoundSwitch != 1) {
                Response.Write("<a href='Reports.asp?RoundSwitch=1&UnitSwitch=" + UnitSwitch + "&Program=" + Program + "&ChoseSample=" + ChoseSample + "&ChoseLab=" + ChoseLab + "&NavSwitch=5' target='_self'>Click to View Nonrounded Stats.</a>");
            }
            else {
                Response.Write("<a href='Reports.asp?RoundSwitch=0&UnitSwitch=" + UnitSwitch + "&Program=" + Program + "&ChoseSample=" + ChoseSample + "&ChoseLab=" + ChoseLab + "&NavSwitch=5' target='_self'>Click to View Rounded Stats.</a>");
            }
            Response.Write("</div>")//No Print
            Response.Write("</caption>");
            Response.Write("<thead>");
            Response.Write("<td align='center' width='25%'>" + "TEST TITLE</td>");
            Response.Write("<td align='center' width='5%'>" + "&nbsp;</td>");
            Response.Write("<td colspan='2' align='center' width='15%'>" + "LAB DATA</td>");
            Response.Write("<td colspan='2' align='center' width='15%'>" + "AVERAGES</td>");
            Response.Write("<td colspan='2' align='center' width='15%'>" + "STAND. DEVS.</td>");
            Response.Write("<td colspan='2' align='center' width='15%'>" + "RATINGS</td>");
            Response.Write("</tr>");
            Response.Write("</caption>");
            Response.Write("</thead>");
            Response.Write("<tr><td align='center' width='35%'>" + "<div id = 'NoPrint'><font color='#800000' >Click below to see Performance Charts</font></div></td>");
            Response.Write("<td align='center' width='5%'>" + "Units</td>");

            for (j = 1; j < 5; j++) {
                Response.Write("<td align='center' width='8%'>" + ChoseSample + "</td>");
                Response.Write("<td align='center' width='8%'>" + EvenSample + "</td>");
            }
            Response.Write("</tr>");
            Response.Write("</thead>");

            //Start of code that displays the test and labs stats and ratings					
            var Duplicate = false
            while (!LabRec.EOF) {
                if (ChoseLab == LabRec.Fields('Labnum').Value && ChoseSample == LabRec.Fields('OddSample').Value && Duplicate == false) {
                    var Remarks = LabRec.Fields('Remarks').Value
                    Duplicate = true
                    var SampleStats = TheComm2.Execute(adCmdText + adExecuteNoRecords);
                    while (!SampleStats.EOF) {
                        //ProgCount = ProgInfo.Fields.Count;
                        for (k in ProgTest) {
                            if (ProgTest[k]['TestNumber'] == SampleStats.Fields('TestNumber')) {
                                // Begining of crap code that determines decimals.
                                //  this if is used to determine decimal plaices for a labs test with no data
                                if (LabRec.Fields('Test' + SampleStats.Fields('TestNumber') + 'Odd') != '--' &&
                                    LabRec.Fields('Test' + SampleStats.Fields('TestNumber') + 'Odd') != '@@' &&
                                    LabRec.Fields('Test' + SampleStats.Fields('TestNumber') + 'Odd') != '##' &&
                                    LabRec.Fields('Test' + SampleStats.Fields('TestNumber') + 'Odd') != '%%') {
                                    var OddDec = DecimalPlaces(LabRec.Fields('Test' + SampleStats.Fields('TestNumber') + 'Odd').Value.split(' ').join(''));
                                }
                                else if (!DecRec == null) {
                                    while (DecRec.Fields('Test' + SampleStats.Fields('TestNumber') + 'Odd') == '--' ||
                                        DecRec.Fields('Test' + SampleStats.Fields('TestNumber') + 'Odd') == '@@' ||
                                        DecRec.Fields('Test' + SampleStats.Fields('TestNumber') + 'Odd') == '##' ||
                                        DecRec.Fields('Test' + SampleStats.Fields('TestNumber') + 'Odd') == '%%') { DecRec.MoveNext }
                                    var OddDec = DecimalPlaces(DecRec.Fields('Test' + SampleStats.Fields('TestNumber') + 'Odd'));

                                }

                                if (LabRec.Fields('Test' + SampleStats.Fields('TestNumber') + 'Eve') != '--' &&
                                    LabRec.Fields('Test' + SampleStats.Fields('TestNumber') + 'Eve') != '@@' &&
                                    LabRec.Fields('Test' + SampleStats.Fields('TestNumber') + 'Eve') != '##' &&
                                    LabRec.Fields('Test' + SampleStats.Fields('TestNumber') + 'Eve') != '%%') {
                                    var EveDec = DecimalPlaces(LabRec.Fields('Test' + SampleStats.Fields('TestNumber') + 'Eve').Value.split(' ').join(''));

                                }
                                else if (!DecRec == null) {
                                    while (DecRec.Fields('Test' + SampleStats.Fields('TestNumber') + 'Eve') == '--' ||
                                        DecRec.Fields('Test' + SampleStats.Fields('TestNumber') + 'Eve') == '@@' ||
                                        DecRec.Fields('Test' + SampleStats.Fields('TestNumber') + 'Eve') == '##' ||
                                        DecRec.Fields('Test' + SampleStats.Fields('TestNumber') + 'Eve') == '%%') { DecRec.MoveNext }
                                    var EveDec = DecimalPlaces(DecRec.Fields('Test' + SampleStats.Fields('TestNumber') + 'Eve'));

                                }
                                if (EveDec > OddDec) { var Dec = EveDec }
                                else { var Dec = OddDec }
                                if (Dec == 0) { var Dec = 0 }
                                if (Dec == null) { var Dec = 2 }

                                Response.Write("<tr>");
                                var ChartSwitch = 'PerformanceChart'
                                Response.Write("<td align='center' width='25%'><a href='Charts.html?Program=" + Program + "&SampleNumber=" + ChoseSample + "&LabNumber=" + LabRec.Fields('Labnum').Value + "&NavSwitch=" + ChartSwitch + "&TestNumber=" + SampleStats.Fields('TestNumber') + "' target='_self'>" + ProgTest[k]['TestTitle'] + "</a></td>");

                                // Used to switch to and from SI Units
                                var Unit = ProgTest[0]['TestUnit']
                                var Converter = 1
                                var LabDataOdd = LabRec.Fields('Test' + SampleStats.Fields('TestNumber') + 'Odd').Value.split(' ').join('')
                                var LabDataEven = LabRec.Fields('Test' + SampleStats.Fields('TestNumber') + 'Eve').Value.split(' ').join('')
                                var OddAvg = SampleStats.Fields('OddAvg').Value
                                var EvenAvg = SampleStats.Fields('EvenAvg').Value
                                var OddSD = SampleStats.Fields('OddSD').Value
                                var EvenSD = SampleStats.Fields('EvenSD').Value
                                if (Unit == 'inch' && UnitSwitch == 1) {
                                    Unit = 'cm'
                                    Converter = 2.54
                                }
                                else if (Unit == 'in<sup>2</sup>' && UnitSwitch == 1) {
                                    Unit = 'cm<sup>2</sup>'
                                    Converter = 6.4516004163
                                }
                                else if (Unit == 'lb' && UnitSwitch == 1) {
                                    Unit = 'kg'
                                    Converter = 0.4535927037
                                }
                                else if (Unit == 'lbf' && UnitSwitch == 1) {
                                    Unit = 'kN'
                                    Converter = 4.4482216153
                                }
                                else if (Unit == 'psi' && UnitSwitch == 1) {
                                    Unit = 'MPa'
                                    Converter = 0.0068947591
                                }
                                else if (Unit == 'lb/ft' && UnitSwitch == 1) {
                                    Unit = 'kg/m'
                                    Converter = 1.488165038
                                }
                                else if (Unit == 'lb/ft<sup>3</sup>' && UnitSwitch == 1) {
                                    Unit = 'kg/m<sup>3</sup>'
                                    Converter = 16.018463373
                                }
                                else if (Unit == '<sup>o</sup> F' && UnitSwitch == 1) {
                                    Unit = '<sup>o</sup> C'
                                    Converter = 0.5555555556
                                    var LabDataOdd = LabDataOdd - 32
                                    var LabDataEven = LabDataEven - 32
                                    var OddAvg = OddAvg - 32
                                    var EvenAvg = EvenAvg - 32
                                }
                                if (UnitSwitch == 1) {
                                    LabDataOdd = parseFloat(RoundASTM(LabDataOdd * Converter, Dec)).toFixed(Dec);
                                    LabDataEven = parseFloat(RoundASTM(LabDataEven * Converter, Dec)).toFixed(Dec);
                                }
                                Response.Write("<td align='center' width='5%'>" + Unit + "</td>");
                                Response.Write("<td align='center' width='8%'>" + LabDataOdd + "</td>");
                                Response.Write("<td align='center' width='8%'>" + LabDataEven + "</td>");
                                if (RoundSwitch == 1) {
                                    if (OddAvg == 0) { Response.Write("<td align='center' width='8%'> -- </td>"); }
                                    else { Response.Write("<td align='center' width='8%'>" + parseFloat(RoundASTM(OddAvg * Converter, Dec + 5)).toFixed(Dec + 5) + "</td>"); }
                                    if (EvenAvg == 0) { Response.Write("<td align='center' width='8%'> -- </td>"); }
                                    else { Response.Write("<td align='center' width='8%'>" + parseFloat(RoundASTM(EvenAvg * Converter, Dec + 5)).toFixed(Dec + 5) + "</td>"); }
                                    if (OddSD == 0) { Response.Write("<td align='center' width='8%'> -- </td>"); }
                                    else { Response.Write("<td align='center' width='8%'>" + parseFloat(RoundASTM(OddSD * Converter, Dec + 5)).toFixed(Dec + 5) + "</td>"); }
                                    if (EvenSD == 0) { Response.Write("<td align='center' width='8%'> -- </td>"); }
                                    else { Response.Write("<td align='center' width='8%'>" + parseFloat(RoundASTM(EvenSD * Converter, Dec + 5)).toFixed(Dec + 5) + "</td>"); }
                                }
                                else {
                                    if (OddAvg == 0) { Response.Write("<td align='center' width='8%'> -- </td>"); }
                                    else { Response.Write("<td align='center' width='8%'>" + parseFloat(RoundASTM(OddAvg * Converter, Dec)).toFixed(Dec) + "</td>"); }
                                    if (EvenAvg == 0) { Response.Write("<td align='center' width='8%'> -- </td>"); }
                                    else { Response.Write("<td align='center' width='8%'>" + parseFloat(RoundASTM(EvenAvg * Converter, Dec)).toFixed(Dec) + "</td>"); }
                                    if (OddSD == 0) { Response.Write("<td align='center' width='8%'> -- </td>"); }
                                    else { Response.Write("<td align='center' width='8%'>" + parseFloat(RoundASTM(OddSD * Converter, Dec)).toFixed(Dec) + "</td>"); }
                                    if (EvenSD == 0) { Response.Write("<td align='center' width='8%'> -- </td>"); }
                                    else { Response.Write("<td align='center' width='8%'>" + parseFloat(RoundASTM(EvenSD * Converter, Dec)).toFixed(Dec) + "</td>"); }
                                }

                                // Do the math to determine labs ratings
                                var OdddRatings
                                var OddSign = ''
                                var EvenRatings
                                var EvenSign = ''
                                var TempOdd = LabRec.Fields('Test' + SampleStats.Fields('TestNumber') + 'Odd').Value - SampleStats.Fields('OddAvg').Value;
                                var TempEven = LabRec.Fields('Test' + SampleStats.Fields('TestNumber') + 'Eve').Value - SampleStats.Fields('EvenAvg').Value;
                                var Display = SampleStats.Fields('Display') + ''

                                if (Math.abs(TempOdd) <= SampleStats.Fields('OddSD').Value) { OddRatings = '5' }
                                else if (Math.abs(TempOdd) <= SampleStats.Fields('OddSD').Value * 1.5) { OddRatings = '4' }
                                else if (Math.abs(TempOdd) <= SampleStats.Fields('OddSD').Value * 2) { OddRatings = '3' }
                                else if (Math.abs(TempOdd) <= SampleStats.Fields('OddSD').Value * 2.5) { OddRatings = '2' }
                                else if (Math.abs(TempOdd) <= SampleStats.Fields('OddSD').Value * 3) { OddRatings = '1' }
                                else if (Math.abs(TempOdd) > SampleStats.Fields('OddSD').Value * 3) { OddRatings = '0' }
                                else { OddRatings = '--' }
                                if (TempOdd < 0 && OddRatings != 0) { OddSign = '-' };
                                if (Display.substr(0, 1) == 'N') { OddRatings = '**'; OddSign = ''; }


                                if (Math.abs(TempEven) <= SampleStats.Fields('EvenSD').Value) { EvenRatings = '5' }
                                else if (Math.abs(TempEven) <= SampleStats.Fields('EvenSD').Value * 1.5) { EvenRatings = '4' }
                                else if (Math.abs(TempEven) <= SampleStats.Fields('EvenSD').Value * 2) { EvenRatings = '3' }
                                else if (Math.abs(TempEven) <= SampleStats.Fields('EvenSD').Value * 2.5) { EvenRatings = '2' }
                                else if (Math.abs(TempEven) <= SampleStats.Fields('EvenSD').Value * 3) { EvenRatings = '1' }
                                else if (Math.abs(TempEven) > SampleStats.Fields('EvenSD').Value * 3) { EvenRatings = '0' }
                                else { EvenRatings = '--' }
                                if (TempEven < 0 && EvenRatings != 0) { EvenSign = '-' };
                                if (Display.substr(1, 1) == 'N') { EvenRatings = '**'; EvenSign = ''; }

                                Response.Write("<td align='center' width='8%'>" + OddSign + OddRatings + "</td>");
                                Response.Write("<td align='center' width='8%'>" + EvenSign + EvenRatings + "</td>");
                                Response.Write('</tr>');
                            }
                        }
                        SampleStats.MoveNext;
                    }
                }
                LabRec.MoveNext;
            }

            Response.Write("</table>");
            Response.Write("<h4>");
            Response.Write("The following table details the relationship between the ratings and the averages.");
            Response.Write("<br>")
            Response.Write("More information can be found on our <a href='../Psp/PspStats.html'>statistics page.</a>");
            Response.Write("</h4>");
            Response.Write("<table class='table table-condensed'>");
            Response.Write("<tr>");
            Response.Write("<td>");
            Response.Write("<h5>")
            Response.Write("Ratings (&plusmn 5 is Best, &plusmn 1 is Worst)</td><td align=center>5</td><td align=center>4</td><td align=center>3</td><td align=center>2</td><td align=center>1");
            Response.Write("</h5>");
            Response.Write("</td>");
            Response.Write("</tr>");

            Response.Write("<tr>");
            Response.Write("<td>");
            Response.Write("<h5>")
            Response.Write("Range (Number of S.D.s)<td align=center>Less than 1</td><td align=center>1 to 1.5</td><td align=center>1.5 to 2</td><td align=center>2 to 2.5</td><td align=center>More than 2.5");
            Response.Write("</h5>");
            Response.Write("</td>");
            Response.Write("</tr>");

            Response.Write("<tr>");
            Response.Write("<td>");
            Response.Write("<h5>")
            Response.Write("Number of Labs Per 100 achieving this rating<td align=center>69</td><td align=center>18</td><td align=center>9</td><td align=center>3</td><td align=center>1</td></tr> ");
            Response.Write("</h5>");
            Response.Write("</td>");
            Response.Write("</tr>");


            Response.Write("</table>");

            Response.Write("<hr>");
            Response.Write("<h4>");
            Response.Write("Notes");
            Response.Write("</h4>");



            Response.Write("<ol>");
            Response.Write("<li id='ref-1'><b>- -</b>&nbsp; No data or invalid data on one or both samples.</li>");
            Response.Write("<li id='ref-2'>A rating of zero (0) indicates the data was excluded due to a deviation of three or more <b>S.D.</b> on one or both samples.</li>");
            Response.Write("<li id='ref-3'><b>* *&nbsp;</b>No ratings assigned for any laboratory for this test.</li>");
            if (Remarks != '' && Remarks != null) {
                Response.Write("<li id='ref-4'>&nbsp;" + Remarks + "</li>");
            }
            Response.Write("</ol>");
            Response.Write("<hr>");
            Response.Write("</div>"); //Container
            Response.Write("<div id = 'NoPrint'><div id='includedFooter'></div></div>");
            Response.Write("</body></html>");
            // end of code that generates the html code
            // close the connection and release it
            TheComm2 = null;
            LabRec.close();
            LabRec = null;
            return ok;
            return SelectLab;
        }
        // End Of fRateings function 

        // Calles fRatings function 
        if (!fRatings(Program, Session('Labnum'))) {
            // if lab and pin don't match redirct
            Session("Error") = 'NotIn';
            Response.Redirect("../../DataEntry/ProgSelect.asp?ErrorProg=" + ErrorProg);
        }
    }
    //********************************************************************************************************************************
    if (NavSwitch > 5 && NavSwitch < 10)// Displays summery of stats. 
    {
        //Sets the SQL connection strings
        var SelectAllData = "SELECT OddSample From " + Program + "Rate";
        // Populates the to arrays with the fields data from the selected records. 
        var AllData = TheConn.Execute(SelectAllData);
        // Creat the Navagation pod based on Sample Number.				
        AllData.MoveFirst;
        while (!AllData.EOF) {
            SampNos[SampCount] = AllData('OddSample').Value
            SampCount = SampCount + 1
            AllData.MoveNext;
        }
        SampNos.sort(function (a, b) { return a - b; });
        for (p = 1; p <= SampCount; p++) {
            if (SampNos[p - 1] != SampNos[p]) {
                SampList[SampListCount] = SampNos[p - 1];
                SampListCount = SampListCount + 1
                Response.Write(SampList[SampListCount])
            }
        }
        SampListCount = SampListCount - 1
        SampList.sort(function (a, b) { return a - b; });
        if (NavSwitch == 7) {
            if (ChoseSample == SampList[0]) {
                ChoseSample = SampList[SampListCount]
            }
            else {
                ChoseSample = +ChoseSample - 2
            }
        }
        if (NavSwitch == 8) {
            if (ChoseSample == SampList[SampListCount]) {
                ChoseSample = SampList[0]
            }
            else {
                ChoseSample = +ChoseSample + 2
            }
        }
        else if (NavSwitch == 9) {
            ChoseSample = Request('ChoseSample')
        }
        var EvenSample = +ChoseSample
        if (EvenSample == 'NaN') {
            EvenSample = SampList[0]
        }
        EvenSample = ++EvenSample

        Response.Write("<div id='NoPrint'>");
        Response.Write("<div class='float-right'>");
        Response.Write("<div class='form-group'>");
        Response.Write("<a href='../Psp/Reports.html'><button type='submit' class='btn btn-primary'>Back to Reports Page</button></a>");
        Response.Write("</div>");
        Response.Write("</div>");


        Response.Write("<table border='0' align='center'>");

        Response.Write("<tr>");
        Response.Write("<td>");
        Response.Write("<a href='Reports.asp?Program=" + Program + "&ChoseSample=" + ChoseSample + "&NavSwitch=" + 7 + "'target='_self'>");
        Response.Write("<span class='icon-backward2'>");
        Response.Write("</span></a>");
        Response.Write("</td>");

        Response.Write("<td align='center'>");
        Response.Write("<a href='javascript:window.print()'target='_self'>Print Page</a>");
        Response.Write("<br>");
        Response.Write("<form name='RatingsForm' METHOD='POST' ACTION='Reports.asp?Program=" + Program + "&NavSwitch=" + 6 + "'>");
        Response.Write("<select language='javascript' onChange='this.form.submit()' name='ChoseSample' >");
        Response.Write("<option selected  value='" + ChoseSample + "'>" + ChoseSample + " & " + EvenSample + '' + "</option>");
        for (j = 0; j < SampList.length; j++) Response.Write("<option  value='" + SampList[j] + "'>" + SampList[j] + " & " + ++SampList[j] + '' + "</option>");
        Response.Write("</select>");
        Response.Write("</form>");
        Response.Write("</td>");

        Response.Write("<td>");
        Response.Write("<a href='Reports.asp?Program=" + Program + "&ChoseSample=" + ChoseSample + "&NavSwitch=" + 8 + "'target='_self'>");
        Response.Write("<span class='icon-forward3'>");
        Response.Write("</span></a>");
        Response.Write("</td>");
        Response.Write("</tr>");

        Response.Write("</table>");

        Response.Write("</div>");// No Print

        var SelectRate = "DECLARE @ChoseSample AS Int;"
        SelectRate += "SELECT @ChoseSample = ?;"
        SelectRate += "SELECT * FROM " + Program + "Rate WHERE (OddSample = @ChoseSample);"

        var TheComm = Server.CreateObject("ADODB.Command")
        TheComm.ActiveConnection = TheConn
        TheComm.CommandText = SelectRate;
        TheComm.Parameters.Append(TheComm.CreateParameter('@ChoseSample', 3, 1, 3, ChoseSample));

        var Labs = TheComm.Execute(adCmdText + adExecuteNoRecords);

        var SelectTestStats = "DECLARE @ChoseSample AS Int;"
        SelectTestStats += "SELECT @ChoseSample = ?;"
        SelectTestStats += "SELECT * FROM " + Program + "Stats WHERE (OddSample = @ChoseSample);"

        var TheComm2 = Server.CreateObject("ADODB.Command")
        TheComm2.ActiveConnection = TheConn
        TheComm2.CommandText = SelectTestStats;
        TheComm2.Parameters.Append(TheComm2.CreateParameter('@ChoseSample', 3, 1, 3, ChoseSample));

        var TestStats = TheComm2.Execute(adCmdText + adExecuteNoRecords);


        Response.Write("<h3>" + "Report Date: " + TestStats.Fields('FinalRptDate') + "</h3>");


        Response.Write("<h3>cclr Proficiency Sample Summery of Statistics</h3>");
        if (Program == 'Concrete' || Program == 'Rebar' || Program == 'MasMort' || Program == 'MasCem') {
            Response.Write("<h3>" + "cclr " + ProgInfo.Fields('Material').Value + " Proficiency Sample Nos. " + ChoseSample + " & " + EvenSample + "</h3>");
        }
        else {
            Response.Write("<h3>" + "cclr " + ProgInfo.Fields('Material').Value + " " + ProgInfo.Fields('Sample Type').Value + " Proficiency Sample Nos. " + ChoseSample + " & " + EvenSample + "</h3>");
        }



        Response.Write("<table class='table table-bordered table-striped'>");

        Response.Write("<thead>");

        Response.Write("<tr>");
        Response.Write("<td align='center' width='330'>" + "TEST TITLE</td>");
        Response.Write("<td></td>");
        Response.Write("<td></td>");
        Response.Write("<td colspan='3' align='center'><b><u>Sample: " + ChoseSample + "</u></b></td>");
        Response.Write("<td colspan='3' align='center'><b><u>Sample: " + EvenSample + "</u></b></td>");

        Response.Write("</tr>");

        Response.Write("<tr>");
        Response.Write("<td align='center' width='30%'>" + "<div id = 'NoPrint'><font color='#800000' >Click below to see Scatter Diagrams.</font></div></td>");
        Response.Write("<td align='center'>" + "UNIT</td>");
        Response.Write("<td align='center' width='10%''>" + "# of Labs</td>");
        Response.Write("<td align='center'>" + "Average</td>");
        Response.Write("<td align='center'>" + "S.D.</td>");
        Response.Write("<td align='center'>" + "C.V.</td>");
        Response.Write("<td align='center'>" + "Average</td>");
        Response.Write("<td align='center'>" + "S.D.</td>");
        Response.Write("<td align='center'>" + "C.V.</td>");
        Response.Write("</tr>");

        Response.Write("</thead>");

        var TestStats = TheComm2.Execute(adCmdText + adExecuteNoRecords);
        while (!TestStats.EOF) {
            ProgCount = ProgInfo.Fields.Count;
            for (j = 5; j < ProgCount; j = j + 3) {
                if (TestStats.Fields('TestNumber') == ProgInfo.Fields(j).Value) {
                    //Labs.MoveFirst;
                    var Labs = TheComm.Execute(adCmdText + adExecuteNoRecords);
                    var Dec = DecimalPlaces(Labs.Fields('Test' + TestStats.Fields('TestNumber') + 'Odd').Value);
                    var NumLabs = 0
                    while (!Labs.EOF) {
                        if (Labs.Fields('Test' + TestStats.Fields('TestNumber') + 'Odd').Value != '--' || Labs.Fields('Test' + TestStats.Fields('TestNumber') + 'Eve').Value != '--') {
                            NumLabs = NumLabs + 1
                        }
                        Labs.MoveNext;
                    }
                    Response.Write("<tr>");
                    Response.Write("<td align='center' width='330'><a href='Charts.html?Program=" + Program + "&ChoseSample=" + ChoseSample + "&ChoseLab=-1&NavSwitch=12&TestNumber=" + TestStats.Fields('TestNumber') + "' target='_self'>" + ProgInfo.Fields(j - 2).Value + "</a></td>");
                    Response.Write("<td align='center' width='110'>" + ProgInfo.Fields(j - 1).Value + "</td>");
                    Response.Write("<td align='center' width='110'>" + NumLabs + "</td>");
                    Response.Write("<td align='center' width='110'>" + parseFloat(RoundASTM(TestStats.Fields('OddAvg').Value, Dec)).toFixed(Dec) + "</td>");
                    Response.Write("<td align='center' width='110'>" + parseFloat(RoundASTM(TestStats.Fields('OddSD').Value, Dec + 1)).toFixed(Dec) + "</td>");
                    var OddCV = 100 * (TestStats.Fields('OddSD').Value / TestStats.Fields('OddAvg').Value)
                    Response.Write("<td align='center' width='110'>" + parseFloat(RoundASTM(OddCV, Dec)).toFixed(Dec) + "</td>");
                    Response.Write("<td align='center' width='110'>" + parseFloat(RoundASTM(TestStats.Fields('EvenAvg').Value, Dec)).toFixed(Dec) + "</td>");
                    Response.Write("<td align='center' width='110'>" + parseFloat(RoundASTM(TestStats.Fields('EvenSD').Value, Dec + 1)).toFixed(Dec) + "</td>");
                    var EvenCV = 100 * (TestStats.Fields('EvenSD').Value / TestStats.Fields('EvenAvg').Value)
                    Response.Write("<td align='center' width='110'>" + parseFloat(RoundASTM(EvenCV, Dec)).toFixed(Dec) + "</td>");
                    Response.Write("</tr>");
                }
            }
            TestStats.MoveNext;
        }
        Response.Write("</table>");
        Response.Write("<h4>");
        Response.Write("More information can be found on our <a href='../Psp/PspStats.html'>statistics page.</a>");
        Response.Write("</h4>");
        Response.Write("</div>"); //Container
        Response.Write("<div id = 'NoPrint'><div id='includedFooter'></div></div>");
        Response.Write("</body></html>");
    }
    TheComm2 = null;
    //********************************************************************************************************************************
    if (NavSwitch == 10)// Displays preliminary report
    {
        //boots smarties out to home if data collection open!!
        var NewDate = new Date()
        if (NewDate < ProgInfo('CloseDate')) {
            Response.Redirect('ProgSelect.asp?NavSwitch=02')
        }
        //  Querry "Data" Data base to get the current Odd Sample No.
        // Must use Request('Program') for Cmuabs and CmuComp to work
        var SelectOddSample = "SELECT OddSample From " + Program + "Data";
        var OddSample = TheConn.Execute(SelectOddSample);
        Labs.MoveFirst;
        var OddSample = OddSample.Fields(0).Value
        var EvenSample = OddSample;
        EvenSample = ++EvenSample;


        Response.Write("<div id = 'NoPrint'>");
        Response.Write("<div class='float-right'>");
        Response.Write("<a href='ProgSelect.asp?NavSwitch=10'><button type='submit' class='btn btn-primary'>Program Selection Page</button></a>");
        Response.Write("</div>");
        Response.Write("<table align=center><tr><td><h3><a href='javascript:window.print()'target='_self'>Print Page</a></h3></td></tr></table>");
        Response.Write("</div>");

        //Response.Write("<h3>" + "Report Date: " + TestStats.Fields('FinalRptDate') + "</h3>");
        Response.Write("<h3>cclr Proficiency Sample Preliminary Report</h3>");
        if (Program == 'Concrete' || Program == 'Rebar' || Program == 'MasMort' || Program == 'MasCem') {
            Response.Write("<h3>" + "cclr " + ProgInfo.Fields('Material').Value + " Proficiency Sample Nos. " + OddSample + " & " + EvenSample + "</h3>");
        }
        else {
            Response.Write("<h3>" + "cclr " + ProgInfo.Fields('Material').Value + " " + ProgInfo.Fields('Sample Type').Value + " Proficiency Sample Nos. " + OddSample + " & " + EvenSample + "</h3>");
        }


        Response.Write("<table class='table table-bordered table-striped'>");
        Response.Write("<thead>");
        Response.Write("<tr>");
        Response.Write("<td align='center'>" + "TEST TITLE</td>");
        Response.Write("<td></td>");
        Response.Write("<td></td>");
        Response.Write("<td colspan='3' align='center'><b><u>Sample: " + OddSample + "</u></b></td>");
        Response.Write("<td colspan='3' align='center'><b><u>Sample: " + EvenSample + "</u></b></td>");
        Response.Write("</tr>");
        Response.Write("<tr>");
        Response.Write("<td></td>");
        Response.Write("<td align='center'>" + "UNIT</td>");
        Response.Write("<td align='center' width='10%''>" + "# of Labs</td>");
        Response.Write("<td align='center'>" + "Average</td>");
        Response.Write("<td align='center'>" + "S.D.</td>");
        Response.Write("<td align='center'>" + "C.V.</td>");
        Response.Write("<td align='center'>" + "Average</td>");
        Response.Write("<td align='center'>" + "S.D.</td>");
        Response.Write("<td align='center'>" + "C.V.</td>");
        Response.Write("</tr>");
        Response.Write("</thead>");
        Response.Write("<tbody>");

        var TestNum = [];
        var TestUnit = [];
        var TestName = [];
        var OutLiers = [];
        for (j = 8; j <= ProgInfo.Fields.Count; j = j + 3) {
            if (ProgInfo.Fields(j).Value != '') {
                TestNum.push(ProgInfo.Fields(j))
                TestUnit.push(ProgInfo.Fields(j - 1))
                TestName.push(ProgInfo.Fields(j - 2))
                var SelectData = "SELECT Labnum,Test" + ProgInfo.Fields(j) + "Odd,Test" + ProgInfo.Fields(j) + "Eve From " + Program + "Data";
                var Labs = TheConn.Execute(SelectData);
                var LabNum = [];
                var OddData = [];
                var EvenData = [];
                var OutCount = 0
                Labs.MoveFirst;
                while (!Labs.EOF) {
                    if (
                        Labs.Fields('Test' + ProgInfo.Fields(j).Value + 'Odd').Value != '' && Labs.Fields('Test' + ProgInfo.Fields(j).Value + 'Eve').Value != ''
                        && Labs.Fields('Test' + ProgInfo.Fields(j).Value + 'Odd').Value != ' ' && Labs.Fields('Test' + ProgInfo.Fields(j).Value + 'Eve').Value != ' '
                        && Labs.Fields('Test' + ProgInfo.Fields(j).Value + 'Odd').Value != null && Labs.Fields('Test' + ProgInfo.Fields(j).Value + 'Eve').Value != null
                        && Labs.Fields('Test' + ProgInfo.Fields(j).Value + 'Odd').Value != 'BadData' && Labs.Fields('Test' + ProgInfo.Fields(j).Value + 'Eve').Value != 'BadData'
                        && Labs.Fields('Test' + ProgInfo.Fields(j).Value + 'Odd').Value != 'undefined' && Labs.Fields('Test' + ProgInfo.Fields(j).Value + 'Eve').Value != 'undefined'
                    ) {
                        LabNum.push([OutCount], [(Labs.Fields('Labnum').Value)])
                        OddData.push(Labs.Fields('Test' + ProgInfo.Fields(j).Value + 'Odd').Value)
                        EvenData.push(Labs.Fields('Test' + ProgInfo.Fields(j).Value + 'Eve').Value)
                    }
                    OutCount++
                    Labs.Move(4);
                }


                var OddStats = (Stats(OddData));
                var EvenStats = (Stats(EvenData));
                var Dec = DecimalPlaces(OddData[0]);

                // creates copyies of the arrays to put throught the "OutLier removal process"
                var LabNum3SD = LabNum.slice();
                var OddData3SD = OddData.slice();
                var EvenData3SD = EvenData.slice();
                // Calculate the stats to use to Run through the Loop and remove the array elements 3 + standard deviations from mean.
                var OddStats3SD = (Stats(OddData3SD));
                var EvenStats3SD = (Stats(EvenData3SD));

                var Run = 1
                var Loop = 0
                while (Loop < Run) {
                    var OutTemp = ''
                    for (i in OddData3SD) {
                        if (Math.abs(OddData3SD[i]) > 3 * OddStats3SD.SD + OddStats3SD.Avg || Math.abs(EvenData3SD[i]) > 3 * EvenStats3SD.SD + EvenStats3SD.Avg) {
                            OutTemp = OutTemp + ', ' + LabNum3SD[i]
                            LabNum3SD.splice(i, 1)
                            OddData3SD.splice(i, 1)
                            EvenData3SD.splice(i, 1)
                        }
                    }
                    OutTemp = OutTemp.substring(2)
                    OutLiers.push(OutTemp)
                    var OddStats3SD = (Stats(OddData3SD));
                    var EvenStats3SD = (Stats(EvenData3SD));
                    Loop++
                }

                Response.Write("<td align='center' width='20%'>" + ProgInfo.Fields(j - 2).Value + "</td>");
                Response.Write("<td align='center' width='10%'>" + ProgInfo.Fields(j - 1).Value + "</td>");
                Response.Write("<td align='center' width='10%'>&nbsp;&nbsp;&nbsp;&nbsp;" + OddData.length + "</td>");
                Response.Write("<td align='center' width='10%'>" + parseFloat(RoundASTM(OddStats.Avg, Dec)).toFixed(Dec) + "</td>");
                Response.Write("<td align='center' width='10%'>" + parseFloat(RoundASTM(OddStats.SD, Dec + 1)).toFixed(Dec) + "</td>");
                Response.Write("<td align='center' width='10%'>" + parseFloat(RoundASTM(OddStats.CV, Dec)).toFixed(Dec) + "</td>");
                Response.Write("<td align='center' width='10%'>" + parseFloat(RoundASTM(EvenStats.Avg, Dec)).toFixed(Dec) + "</td>");
                Response.Write("<td align='center' width='10%'>" + parseFloat(RoundASTM(EvenStats.SD, Dec + 1)).toFixed(Dec) + "</td>");
                Response.Write("<td align='center' width='10%'>" + parseFloat(RoundASTM(EvenStats.CV, Dec)).toFixed(Dec) + "</td></tr>");

            }
        }
        Response.Write("</tbody>")
        Response.Write("</table>")

        Response.Write("<table class='table table-bordered table-striped'>");
        Response.Write("<thead>");
        Response.Write("<h4>");
        Response.Write("* ELIMINATED LABS:  Data over three S.D. from the mean.");
        Response.Write("</h4>");
        Response.Write("</thead>");
        Response.Write("<tbody>")
        for (O = 0; O < TestName.length; O++) {
            if (OutLiers[O] != '') {
                Response.Write("<td align='center' width='20%'>" + TestName[O] + "</td>");
                Response.Write("<td align='left' width='80%'>" + OutLiers[O] + "</td></tr>");
            }
        }
        Response.Write("</tbody>")
        Response.Write("</table>");

        Response.Write("<div id = 'NoPrint'>");
        Response.Write("<h4>");
        Response.Write("More information can be found on our <a href='../Psp/PspStats.html'>statistics page.</a>");
        Response.Write("</h4>");
        Response.Write("</div>");

        Response.Write("</div>"); //Container
        Response.Write("<div id = 'NoPrint'><div id='includedFooter'></div></div>");
        Response.Write("</body></html>");
    }
    //********************************************************************************************************************************

    TheComm = null;
    TheComm2 = null;
    Close_Connection();
</script>
<!-- Bootstrap core JavaScript
================================================== -->
<!-- Placed at the end of the document so the pages load faster -->
<script src="../js/jquery/modernizr-2.8.3.min.js"></script>
<script src="https://ajax.googleapis.com/ajax/libs/jquery/1.11.1/jquery.min.js"></script>
<script>window.jQuery || document.write('<script src="../js/jquery/jquery-1.11.1.min.js"><\/script>')</script>
<script src="../js/jquery/bootstrap.min.js"></script>
<script src="../js/jquery/jquery.metadata.js"></script>
<script src="../js/jquery/jquery-ui.custom.min.js"></script>
<!-- IE10 viewport hack for Surface/desktop Windows 8 bug -->
<script src="../js/ie10-viewport-bug-workaround.js"></script>
<!-- jquery scripts to include header and footer files-->
<script>$(function () { $("#includedHeader").load("../Header.html"); });</script>
<script>$(function () { $("#includedFooter").load("../Footer.html"); });</script>
