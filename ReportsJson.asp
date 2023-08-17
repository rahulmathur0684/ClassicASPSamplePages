<%@  language="JScript" %>
<!--#include file="../ConnectionString/Connection.asp" -->
<script language="JavaScript" runat="server"> 
    // Data base connectin objects.   
    var TheAccessConn = Server.CreateObject("ADODB.Connection");// Create a connection object for Access.
    // Used for all three reports Ratings Summmation of stats and prelim report. Data base record sets using parameterized query
    var RecordSet = null;  // Used to store any record set from parameterized query untill the record set is  placed into an array of arrayes.    
    // These variables are used for all three reports, Ratings, Summmation of Stats and Prelim Report.
    var NavSwitch = Request('NavSwitch');
    var ChoseSample = Request('ChoseSample');
    var NumOfTests = 0; // ** Value can only be an integer. Used by old ASP to tranfer the number of tests from the server side of page to the client JavaScript functions.
    var LatestSample = null;
    var TestNum = [];
    var SampNo = 0;
    //Record set Arrays and associated variables
    var CmdTextStr = null; // The command text string for the command object.
    var g = null;// variable used to loop arrays
    var ProgTest = []; // Used for all three reports, Ratings Summmation of Stats and Prelim Report.
    var Program = Request('Program');
    if (Program == 'undefined') {
        CmdTextStr = "SELECT Program, Material, SampleType, CloseDate FROM ProgramInfo Where (DisplayOrder = 1 and Displaystatus='YY');";
        TheComm = Server.CreateObject("ADODB.Command");// Create a command object.
        TheComm.ActiveConnection = TheConn;// Assine the comand object to the connectin object TheConn.
        TheComm.CommandText = CmdTextStr;// Assine the the comand text to the command object TheComm.	
        // Use parameterized query to stop SQL Injection attacks      
        RecordSet = TheComm.Execute(adCmdText + adExecuteNoRecords);// Record set from parameterized query table "ProgramInfo" to retrieve heading info for the selected program.
        TheComm = null;// Release the command object.
        var Now = new Date();
        var CloseData = [];
        g = 0;
        // This while reads in all the programs in ProgramInfo. this will be used for the select box on the client.
        while (!RecordSet.EOF) {
            CloseData[g] = [];
            var CloseDate = new Date(RecordSet.Fields('CloseDate').value);
            if (Now > CloseDate) {
                CloseData[g].Program = RecordSet.Fields('Program').value;
                CloseData[g].CloseDate = RecordSet.Fields('CloseDate').value;
                g++;
            }
            RecordSet.MoveNext;
        }
        RecordSet.close();
        // sort in order of CloseDate for the obvious reason. 
        CloseData.sort(function (a, b) { return b.CloseDate - a.CloseDate; });
        var Program = CloseData[0].Program;
    }

    // Used for all three reports, Ratings, Summmation of Stats and Prelim Report.
    CmdTextStr = 'DECLARE @Program AS Char(10);'
    CmdTextStr += 'SELECT @Program = ?;'
    CmdTextStr += 'SELECT Program, Material, SampleType, CloseDate, TestNumber, TestTitle, DisplayOrder, DisplayStatus, TestUnit, TestDec FROM ProgramInfo Where (Program = @Program);'
    TheComm = Server.CreateObject("ADODB.Command");// Create a command object.
    TheComm.ActiveConnection = TheConn;// Assine the comand object to the connectin object TheConn.
    TheComm.CommandText = CmdTextStr;// Assine the the comand text to the command object TheComm.	
    // Use parameterized query to stop SQL Injection attacks
    TheComm.Parameters.Append(TheComm.CreateParameter('@Program', 129, 1, 10, Program)); // Create the parameter query based on the variable Program. Can be Concrete, Portland Cement ect. 
    RecordSet = TheComm.Execute(adCmdText + adExecuteNoRecords);// Record set from parameterized query table "ProgramInfo" to retrieve heading info for the selected program.
    TheComm = null;// Release the command object.

    //  This while reads in the data from a specific program in ProgInfo table of pspdata data base into an array of arrays "ProgTest".
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
        if (ProgTest[g]['DisplayStatus'].substr(2, 1) != 'N') {
            NumOfTests++; // ** Value can only be an integer. Used by old ASP to tranfer the number of tests from the server side of page to the client JavaScript functions.
        };
        RecordSet.MoveNext;
        g++
    };
    RecordSet.close();
    var Material = ProgTest[0]['Material']
    var SampleType = ProgTest[0]['SampleType']
    if (Session('Labnum') == null) {
        if (Request.Cookies('Labnum') == null) {
            Response.Redirect('../default.asp')
        }
        else {
            Session("Labnum") = Request.Cookies('Labnum') + '';
        };
    };
    // sort in order of DispayOrder for the obvious reason. 
    ProgTest.sort(function (a, b) { return a.DisplayOrder - b.DisplayOrder; });
    // Error message for ?   
    ErrorProg = Material + " " + SampleType;
    if (NavSwitch != 'Prelim')// Used for Summary of Stats and Ratings.
    {
        // Begining of function RoundASTM()-->Rounds to one more decimal place than decimal places displayed.
        function RoundASTM(Number, Decimals) {
            var i = parseFloat(Number);
            if (isNaN(i)) { i = 0.00; }
            var minus = '';
            if (i < 0) { minus = '-'; }
            i = Math.abs(i);
            if (Decimals == -1) { i = parseInt((i + 5) * 0.1); i = i / 0.1; }
            else if (Decimals == 0) { i = parseInt(i + .5); }  // or use next line
            //else if (Decimals ==  0){i = parseInt((i + .5) * 1); i = i / 1;}
            else if (Decimals == 1) { i = parseInt((i + .05) * 10); i = i / 10; }
            else if (Decimals == 2) { i = parseInt((i + .005) * 100); i = i / 100; }
            else if (Decimals == 3) { i = parseInt((i + .0005) * 1000); i = i / 1000; }
            else if (Decimals == 4) { i = parseInt((i + .00005) * 10000); i = i / 10000; }
            else if (Decimals == 5) { i = parseInt((i + .000005) * 100000); i = i / 100000; }
            else if (Decimals == 6) { i = parseInt((i + .0000005) * 1000000); i = i / 1000000; }
            RoundedNum = new String(i);
            RoundedNum = minus + RoundedNum;
            return RoundedNum;
        };
        // End of function RoundASTM()
    };

    //********************************************************************************************************************************
    if (NavSwitch == 'Ratings')// Dispays Ratings Json and Muti Ratings Json.
    {
        var SampleStats = [];
        var RoundSwitch = 0;
        CmdTextStr = "SELECT * FROM " + Program + "Stats Where Outliers != 'All Data' ORDER BY CAST(OddSample AS BIGINT) desc";
        TheComm = Server.CreateObject("ADODB.Command")// Create a command object.
        TheComm.ActiveConnection = TheConn;// Assine the comand object to the connectin object TheConn.
        TheComm.CommandText = CmdTextStr;// Assine the the comand text to the command object TheComm.      
        RecordSet = TheComm.Execute(adCmdText + adExecuteNoRecords);
        TheComm = null;// Release the command object.

        while (!RecordSet.EOF) {
            var Test = RecordSet.Fields('TestNumber').value.split(' ').join('');
            var OddSamp = RecordSet.Fields('OddSample').value.split(' ').join('');
            var Outliers = RecordSet.Fields('Outliers').value + ''
            if (SampNo != OddSamp) {
                SampleStats[OddSamp] = [];
                TestNum[OddSamp] = [];
            }
            TestNum[OddSamp][Test] = [];
            TestNum[OddSamp][Test]['TestNumber'] = Test;
            for (h in ProgTest) {
                if (ProgTest[h]['TestNumber'].split(' ').join('') == Test) //&& Outliers != 'All Data')
                {
                    TestNum[OddSamp][Test]['TestTitle'] = ProgTest[h]['TestTitle'];
                    TestNum[OddSamp][Test]['TestUnit'] = ProgTest[h]['TestUnit'];
                    TestNum[OddSamp][Test]['TestDec'] = ProgTest[h]['TestDec'];
                    TestNum[OddSamp][Test]['DisplayOrder'] = ProgTest[h]['DisplayOrder'];
                }
            }
            SampNo = OddSamp;
            SampleStats[OddSamp][Test] = [];
            SampleStats[OddSamp][Test]['Display'] = RecordSet.Fields('Display').value;
            SampleStats[OddSamp][Test]['OddSample'] = OddSamp.split(' ').join('');
            SampleStats[OddSamp][Test]['TestNumber'] = Test.split(' ').join('');
            SampleStats[OddSamp][Test]['OddSD'] = RecordSet.Fields('OddSD').value;
            SampleStats[OddSamp][Test]['EvenSD'] = RecordSet.Fields('EvenSD').value;
            SampleStats[OddSamp][Test]['OddAvg'] = RecordSet.Fields('OddAvg').value;
            SampleStats[OddSamp][Test]['EvenAvg'] = RecordSet.Fields('EvenAvg').value;
            SampleStats[OddSamp][Test]['FinalRptDate'] = RecordSet.Fields('FinalRptDate').value;
            LatestSample = SampleStats[OddSamp][Test]['OddSample']
            RecordSet.MoveNext;
        };
        // sort in order of DispayOrder for the obvious reason.         
        RecordSet.close();

        // These variables are used for Ratings Report.
        var RatingsJson = null;
        var RoundSwitch = Request('RoundSwitch');
        var Remarks = null;
        var BufferCount = 0;
        var OddSample = null;
        var EvenSample = null;
        var Dec = null;
        var TestNumber = null;
        var TestTitle = null;
        var Unit = null;
        var LabDataOdd = null;
        var LabDataEven = null;
        var OddAvg = null;
        var EvenAvg = null;
        var OddSD = null;
        var EvenSD = null;
        var OddCV = null;
        var EvenCV = null;
        var FinalRptDate = null;
        var OdddRatings = null;
        var OddSign = null;
        var EvenRatings = null;
        var EvenSign = null;
        var TempOdd = null;
        var TempEven = null;
        var Display = null;
        var MultiLabNum = [];
        var ProgName = [];
        var TestStats = [];
        var Labs = [];
        var LabData = [];
        var LabRec = null;  // Record set from data table with all data from givin program. 
        var FieldCount = null;
        var SingleRec = null;
        var LabRec = [];
        var MultiLabNum = null;
        var EvenSample = 0;
        var TestStats;  // Record set from query table "Prog" States to retrieve all Stats.
        var SampCount = 0;
        var SampNos = []; // Used for Summary of Stats and Ratings
        var SampList = [];
        var DeclareRec = "DECLARE @Labnum AS Char(5)";
        var ParamRec = "SELECT @Labnum = ?";
        var SqlRec = "SELECT * FROM " + Program + "Rate WHERE (Labnum = @Labnum)";

        var TheComm = Server.CreateObject("ADODB.Command");
        TheComm.ActiveConnection = TheConn
        if (Session('Pinnum').length == 5) {
            if (Session('Pinnum') == 51262)// cclr's back door
            {
                CmdTextStr = "DECLARE @LatestSample AS Int;"
                CmdTextStr += "SELECT @LatestSample = ?;"
                CmdTextStr += "SELECT * FROM " + Program + "Rate WHERE (OddSample >= @LatestSample);"
                TheComm.CommandText = CmdTextStr;
                TheComm.Parameters.Append(TheComm.CreateParameter('@LatestSample', 3, 1, 3, LatestSample - 4));
            }
            else {
                CmdTextStr = "DECLARE @ReceiverPin AS Int;"
                CmdTextStr += "SELECT @ReceiverPin = ?;"
                CmdTextStr += "SELECT * FROM MultiLabs WHERE (ReceiverPin = @ReceiverPin);"
                TheComm.CommandText = CmdTextStr;
                TheComm.Parameters.Append(TheComm.CreateParameter('@ReceiverPin', 3, 1, 5, Session('Pinnum')));
                RecordSet = TheComm.Execute(adCmdText + adExecuteNoRecords);

                TheComm = null;// Release the command object.	  
                var TheComm = Server.CreateObject("ADODB.Command")
                TheComm.ActiveConnection = TheConn
                TheComm.Parameters.Append(TheComm.CreateParameter('@Labnum', 129, 1, 4, RecordSet('SenderLabno')));
                Count = 1
                while (!RecordSet.EOF) {
                    DeclareRec += ",@Labnum" + Count + " AS Char(5)"
                    ParamRec += ",@Labnum" + Count + " = ?"
                    SqlRec += " OR (Labnum = @Labnum" + Count + ")"
                    TheComm.Parameters.Append(TheComm.CreateParameter('@Labnum' + Count, 129, 1, 5, RecordSet('SenderLabno')));
                    RecordSet.MoveNext
                    Count++
                }
                if (Session('Pinnum') == 49950) // for big specifiers like AAP only returns 3 most recent sample pairs.
                {
                    DeclareRec += ",@LatestSample AS Int"
                    ParamRec += ",@LatestSample = ?"
                    SqlRec = SqlRec.replace(/WHERE /, 'WHERE ('); // add an extra (.	
                    SqlRec += " ) AND (OddSample >= @LatestSample)"
                    TheComm.Parameters.Append(TheComm.CreateParameter('@LatestSample', 3, 1, 3, LatestSample - 4));
                }
                RecordSet.close();
                DeclareRec += ";"
                ParamRec += ";"
                SqlRec += ";"
                var CmdTextStr = DeclareRec + ParamRec + SqlRec
                //Response.Write(CmdTextStr);
            }
        }
        else {
            var Labnum = Session('Labnum');
            CmdTextStr = "DECLARE @Labnum AS Char(5);"
            CmdTextStr += "SELECT @Labnum = ?;"
            CmdTextStr += ";with TBL AS (SELECT ROW_NUMBER() OVER(Partition BY CAST(OddSample AS BIGINT) ORDER BY CAST(OddSample AS BIGINT)) AS Row#, * FROM " + Program + "Rate WHERE (Labnum = @Labnum)) Select * from TBL where Row#=1 Order by convert (int, OddSample);"
            TheComm.Parameters.Append(TheComm.CreateParameter('@Labnum', 129, 1, 5, Labnum));
        }
        TheComm.CommandText = CmdTextStr;
        RecordSet = TheComm.Execute(adCmdText + adExecuteNoRecords);
        TheComm = null;// Release the command object.
        FieldCount = RecordSet.Fields.Count;
        g = 0;
        TestCount = RecordSet.Fields.Count;
        while (!RecordSet.EOF) {
            LabRec[g] = [];
            LabRec[g]['LabNum'] = RecordSet.Fields('LabNum').Value;
            LabRec[g]['OddSample'] = RecordSet.Fields('OddSample').Value;
            //  this for loads all the differint test data into the array and remove nulls	
            for (j = 3; j < TestCount; j++) {
                if (RecordSet.Fields(j).Value == null) {
                    LabRec[g][RecordSet.Fields(j).Name] = '--'
                }
                else {
                    LabRec[g][RecordSet.Fields(j).Name] = RecordSet.Fields(j).Value;
                }
            };
            RecordSet.MoveNext;
            g++
        }
        RecordSet.close();

        // don't remove need for AASHTO because Taable is so large.
        Response.Buffer = false;
        //Start of code that displays the test and labs stats and ratings Json					
        if (Session('Pinnum').length == 4) {
            RatingsJson = '{"NavSwitch":"' + NavSwitch + '","Program":"' + Program + '","Material":"' + Material + '","SampleType":"' + SampleType + '", "NumOfTests":' + NumOfTests + ', "Labnum":' + Labnum + ',"data":[[';
        }
        else {
            RatingsJson = '{"NavSwitch":"' + NavSwitch + '","Program":"' + Program + '","Material":"' + Material + '","SampleType":"' + SampleType + '", "NumOfTests":' + NumOfTests + ',  "data":[[';
        }
        var FirstLoop = true;
        var BufferCount = 0
        for (h in LabRec) {

            Remarks = LabRec[h]['Remarks']
            if (BufferCount == 1000) {
                Response.Write(RatingsJson);
                RatingsJson = '';
                BufferCount = 0
            }
            BufferCount++

            for (k in TestNum[LabRec[h]['OddSample']]) {
                OddSample = LabRec[h]['OddSample']
                EvenSample = OddSample;
                EvenSample++;

                // Determine the decimals
                Dec = TestNum[LabRec[h]['OddSample']][k]['TestDec'];
                if (Dec == 0.1) { Dec = 1 }
                else if (Dec == 0.01) { Dec = 2 }
                else if (Dec == 0.001) { Dec = 3 }
                else if (Dec == 0.0001) { Dec = 4 }
                else if (Dec == 0.25) { Dec = 2 }
                else if (Dec == 0.5) { Dec = 1 }
                else if (Dec > 1) { Dec = 0 }
                if (RoundSwitch == 1) { Dec = parseInt(Dec) + 5 }

                TestTitle = TestNum[LabRec[h]['OddSample']][k]['TestTitle'];
                TestTitle = TestTitle.replace(/"/g, "")
                Unit = TestNum[LabRec[h]['OddSample']][k]['TestUnit'];
                LabDataOdd = LabRec[h]['Test' + TestNum[LabRec[h]['OddSample']][k]['TestNumber'] + 'Odd'].split(' ').join('');
                LabDataEven = LabRec[h]['Test' + TestNum[LabRec[h]['OddSample']][k]['TestNumber'] + 'Eve'].split(' ').join('');
                OddAvg = SampleStats[LabRec[h]['OddSample']][TestNum[LabRec[h]['OddSample']][k]['TestNumber']]['OddAvg'];
                EvenAvg = SampleStats[LabRec[h]['OddSample']][TestNum[LabRec[h]['OddSample']][k]['TestNumber']]['EvenAvg'];
                OddSD = SampleStats[LabRec[h]['OddSample']][TestNum[LabRec[h]['OddSample']][k]['TestNumber']]['OddSD'];
                EvenSD = SampleStats[LabRec[h]['OddSample']][TestNum[LabRec[h]['OddSample']][k]['TestNumber']]['EvenSD'];
                TestNumber = TestNum[LabRec[h]['OddSample']][k]['TestNumber'];
                FinalRptDate = SampleStats[LabRec[h]['OddSample']][TestNum[LabRec[h]['OddSample']][k]['TestNumber']]['FinalRptDate'];
                var Month = FinalRptDate.substr(0, 3)
                var DayYear = FinalRptDate.slice(-8)
                var Year = FinalRptDate.slice(-4);
                //FinalRptDate = Month + ' ' + DayYear;

                // Do the math to determine labs ratings
                OddSign = ''
                EvenRatings
                EvenSign = ''
                TempOdd = LabRec[h]['Test' + TestNum[LabRec[h]['OddSample']][k]['TestNumber'] + 'Odd'] - SampleStats[LabRec[h]['OddSample']][TestNum[LabRec[h]['OddSample']][k]['TestNumber']]['OddAvg'];
                TempEven = LabRec[h]['Test' + TestNum[LabRec[h]['OddSample']][k]['TestNumber'] + 'Eve'] - SampleStats[LabRec[h]['OddSample']][TestNum[LabRec[h]['OddSample']][k]['TestNumber']]['EvenAvg'];
                Display = SampleStats[LabRec[h]['OddSample']][TestNum[LabRec[h]['OddSample']][k]['TestNumber']]['Display'] + ''

                if (Math.abs(TempOdd) <= SampleStats[LabRec[h]['OddSample']][TestNum[LabRec[h]['OddSample']][k]['TestNumber']]['OddSD']) { OddRatings = '5' }
                else if (Math.abs(TempOdd) <= SampleStats[LabRec[h]['OddSample']][TestNum[LabRec[h]['OddSample']][k]['TestNumber']]['OddSD'] * 1.5) { OddRatings = '4' }
                else if (Math.abs(TempOdd) <= SampleStats[LabRec[h]['OddSample']][TestNum[LabRec[h]['OddSample']][k]['TestNumber']]['OddSD'] * 2) { OddRatings = '3' }
                else if (Math.abs(TempOdd) <= SampleStats[LabRec[h]['OddSample']][TestNum[LabRec[h]['OddSample']][k]['TestNumber']]['OddSD'] * 2.5) { OddRatings = '2' }
                else if (Math.abs(TempOdd) <= SampleStats[LabRec[h]['OddSample']][TestNum[LabRec[h]['OddSample']][k]['TestNumber']]['OddSD'] * 3) { OddRatings = '1' }
                else if (Math.abs(TempOdd) > SampleStats[LabRec[h]['OddSample']][TestNum[LabRec[h]['OddSample']][k]['TestNumber']]['OddSD'] * 3) { OddRatings = '0' }
                else { OddRatings = '--' }
                if (TempOdd < 0 && OddRatings != 0) { OddSign = '-' };
                if (Display.substr(0, 1) == 'N') { OddRatings = '**'; OddSign = '' }

                if (Math.abs(TempEven) <= SampleStats[LabRec[h]['OddSample']][TestNum[LabRec[h]['OddSample']][k]['TestNumber']]['EvenSD']) { EvenRatings = '5' }
                else if (Math.abs(TempEven) <= SampleStats[LabRec[h]['OddSample']][TestNum[LabRec[h]['OddSample']][k]['TestNumber']]['EvenSD'] * 1.5) { EvenRatings = '4' }
                else if (Math.abs(TempEven) <= SampleStats[LabRec[h]['OddSample']][TestNum[LabRec[h]['OddSample']][k]['TestNumber']]['EvenSD'] * 2) { EvenRatings = '3' }
                else if (Math.abs(TempEven) <= SampleStats[LabRec[h]['OddSample']][TestNum[LabRec[h]['OddSample']][k]['TestNumber']]['EvenSD'] * 2.5) { EvenRatings = '2' }
                else if (Math.abs(TempEven) <= SampleStats[LabRec[h]['OddSample']][TestNum[LabRec[h]['OddSample']][k]['TestNumber']]['EvenSD'] * 3) { EvenRatings = '1' }
                else if (Math.abs(TempEven) > SampleStats[LabRec[h]['OddSample']][TestNum[LabRec[h]['OddSample']][k]['TestNumber']]['EvenSD'] * 3) { EvenRatings = '0' }
                else { EvenRatings = '--' }
                if (TempEven < 0 && EvenRatings != 0) { EvenSign = '-' }
                if (Display.substr(1, 1) == 'N') { EvenRatings = '**'; EvenSign = ''; }

                if (FirstLoop != true) {
                    RatingsJson += ',[';
                }
                RatingsJson += '"' + FinalRptDate + '",';
                if (OddSample == "1") { OddSample = "01" };
                if (OddSample == "2") { OddSample = "02" };
                if (OddSample == "3") { OddSample = "03" };
                if (OddSample == "4") { OddSample = "04" };
                if (OddSample == "5") { OddSample = "05" };
                if (OddSample == "6") { OddSample = "06" };
                if (OddSample == "7") { OddSample = "07" };
                if (OddSample == "8") { OddSample = "08" };
                if (OddSample == "9") { OddSample = "09" };

                RatingsJson += '"' + OddSample + '",';

                if (Session('Pinnum').length == 5) {
                    RatingsJson += '"' + LabRec[h]['LabNum'] + '",';
                }
                RatingsJson += '"' + TestTitle + '",';
                RatingsJson += '"' + Unit + '",';
                RatingsJson += '"' + LabDataOdd + '",';
                RatingsJson += '"' + LabDataEven + '",';
                RatingsJson += '"' + OddSign + OddRatings + '",';
                RatingsJson += '"' + EvenSign + EvenRatings + '",';
                if (OddAvg == 0) { RatingsJson += '"--",' }
                else { RatingsJson += '"' + parseFloat(RoundASTM(OddAvg, Dec)).toFixed(parseInt(Dec)) + '",' }
                if (EvenAvg == 0) { RatingsJson += '"--",' }
                else { RatingsJson += '"' + parseFloat(RoundASTM(EvenAvg, Dec)).toFixed(parseInt(Dec)) + '",' }

                // Add one decemal place to the SD if Decimal places are 0 or 4
                if (OddSD == 0) { RatingsJson += '"--",' }
                else if (Dec == 0 || Dec == 4 || RoundSwitch == 1) { RatingsJson += '"' + parseFloat(RoundASTM(OddSD, Dec)).toFixed(parseInt(Dec)) + '",' }
                else { RatingsJson += '"' + parseFloat(RoundASTM(OddSD, Dec + 1)).toFixed(parseInt(Dec) + 1) + '",' }
                if (EvenSD == 0) { RatingsJson += '"--",' }
                else if (Dec == 0 || Dec == 4 || RoundSwitch == 1) { RatingsJson += '"' + parseFloat(RoundASTM(EvenSD, Dec)).toFixed(parseInt(Dec)) + '",' }
                else { RatingsJson += '"' + parseFloat(RoundASTM(EvenSD, Dec + 1)).toFixed(parseInt(Dec) + 1) + '",' }

                RatingsJson += '"' + TestNumber + '"]';
                FirstLoop = false;
            }	// End for (k in SampleStats)
        }	// End for (h in LabRec)

        RatingsJson += ']}';
        Response.Write(RatingsJson);
    }
    //********************************************************************************************************************************
    if (NavSwitch == 'SumOfStats')// Displays summery of stats. 
    {
        // Record set Arrays and associated variables used for Summary of Stats report.
        var TestTitle = null;
        var RoundSwitch = Request('RoundSwitch');
        var Unit = null;
        var FinalRptDate = null;
        var OddSample = null;
        var EvenSample = null;
        var Samples = null;
        var NumLabs = null; // Used to count the number of labs in an sample.
        var Dec = null;  // Used to determine the number of decimal places to display.
        var OddAvg = null;
        var EvenAvg = null;
        var OddSD = null;
        var EvenSD = null;
        var OddCV = null;
        var EvenCV = null;
        var OddSamples = null;  // Record set from ratings table with the odd sample data from the all programs.
        var AllLabsInSamp = []; // Used for Summary of Stats
        var OddSampleNos = [];  // Used for Summary of Stats
        var SumStatsJson = "";
        var SampleStats = [];
        //var RoundSwitch = 0;
        var FirstLoop = true;
        var Now = new Date();
        // Functions
        function GetPrograms() {
            CmdTextStr = "SELECT Program, Material, SampleType, CloseDate FROM ProgramInfo Where (DisplayOrder = 1 and Displaystatus='YY');";
            TheComm = Server.CreateObject("ADODB.Command");// Create a command object.
            TheComm.ActiveConnection = TheConn;// Assine the comand object to the connectin object TheConn.
            TheComm.CommandText = CmdTextStr;// Assine the the comand text to the command object TheComm.	
            // Use parameterized query to stop SQL Injection attacks         
            RecordSet = TheComm.Execute(adCmdText + adExecuteNoRecords);// Record set from parameterized query table "ProgramInfo" to retrieve heading info for the selected program.
            TheComm = null;// Release the command object.
            SumStatsJson += '{"Programs":[';
            var CloseData = [];
            g = 0;
            // This while reads in all the programs in ProgramInfo. this will be used for the select box on the client.
            while (!RecordSet.EOF) {
                CloseData[g] = [];
                SumStatsJson += '"' + RecordSet.Fields('Program').value + '",';
                CloseDate = new Date(RecordSet.Fields('CloseDate').value);
                if (Now > CloseDate) {
                    CloseData[g].Program = RecordSet.Fields('Program').value;
                    CloseData[g].CloseDate = RecordSet.Fields('CloseDate').value;
                    g++;
                }
                RecordSet.MoveNext;
            }
            RecordSet.close();
            SumStatsJson = SumStatsJson.substr(0, SumStatsJson.length - 1); // remove last troublesome ",".
            SumStatsJson += '],';
            // sort in order of CloseDate for the obvious reason. 
            CloseData.sort(function (a, b) { return b.CloseDate - a.CloseDate; });
            var Program = CloseData[0].Program;
            return Program
        }
        GetPrograms()

        CmdTextStr = "SELECT * FROM " + Program + "Rate";
        TheComm = Server.CreateObject("ADODB.Command");// Create a command object.
        TheComm.ActiveConnection = TheConn;// Assine the comand object to the connectin object TheConn.
        TheComm.CommandText = CmdTextStr;// Assine the the comand text to the command object TheComm.
        RecordSet = TheComm.Execute(adCmdText + adExecuteNoRecords);// Record set from query table "Prog" Ratings to retrieve all Data.
        TheComm = null;// Release the command object.
        //  This while reads in the data from a specific program in Prog ratings table of pspdata data base into an array of arrays "AllLabsInSamp".
        g = 0;
        TestCount = RecordSet.Fields.Count;
        g = 0;

        CmdTextStr = "SELECT * FROM " + Program + "Stats";
        TheComm = Server.CreateObject("ADODB.Command")// Create a command object.
        TheComm.ActiveConnection = TheConn;// Assine the comand object to the connectin object TheConn.
        TheComm.CommandText = CmdTextStr;// Assine the the comand text to the command object TheComm.
        RecordSet = TheComm.Execute(adCmdText + adExecuteNoRecords);
        TheComm = null;// Release the command object.

        SumStatsJson += '"NavSwitch":"' + NavSwitch + '","Program":"' + Program + '", "Material":"' + Material + '", "SampleType":"' + SampleType + '", "NumOfTests":' + NumOfTests + ', "data":[[';

        while (!RecordSet.EOF) {
            if (SampNo != RecordSet.Fields('OddSample').value) {
                SampleStats[RecordSet.Fields('OddSample').value] = [];
                TestNum[RecordSet.Fields('OddSample').value] = [];
                OddSampleNos[g] = [];
                OddSampleNos[g]['OddSample'] = RecordSet.Fields('OddSample').value;
                g++
            }
            TestNum[RecordSet.Fields('OddSample').value][RecordSet.Fields('TestNumber').value] = [];
            TestNum[RecordSet.Fields('OddSample').value][RecordSet.Fields('TestNumber').value]['TestNumber'] = RecordSet.Fields('TestNumber').value;

            for (h in ProgTest) {
                if (ProgTest[h]['TestNumber'] == RecordSet.Fields('TestNumber').value.split(' ').join('')) {
                    TestTitle = ProgTest[h]['TestTitle'];
                    Unit = ProgTest[h]['TestUnit'];
                    var Dec = ProgTest[h]['TestDec'];
                }
            }
            SampNo = RecordSet.Fields('OddSample').value;
            if (RecordSet.Fields('NumOfLabs').value == null) { var NumOfLabs = 'View Report'; var Outliers = 'View Report'; } else { NumOfLabs = RecordSet.Fields('NumOfLabs').value; NumOfLabs = NumOfLabs.slice(0, 4); Outliers = RecordSet.Fields('Outliers').value + ''; }

            FinalRptDate = RecordSet.Fields('FinalRptDate').value;
            var Month = FinalRptDate.substr(0, 3)
            var DayYear = FinalRptDate.slice(-8)
            var Year = FinalRptDate.slice(-4);
            //FinalRptDate = Month + ' ' + DayYear;

            OddSample = RecordSet.Fields('OddSample').value;
            EvenSample = OddSample
            EvenSample++;
            Samples = OddSample + ' - ' + EvenSample;
            TestNumber = RecordSet.Fields('TestNumber').value;
            OddAvg = RecordSet.Fields('OddAvg').value;
            EvenAvg = RecordSet.Fields('EvenAvg').value;
            OddSD = RecordSet.Fields('OddSD').value;
            EvenSD = RecordSet.Fields('EvenSD').value;
            OddCV = 100 * (OddSD / OddAvg);
            EvenCV = 100 * (EvenSD / EvenAvg);

            // Retreave and adjust decimal places to display.
            // converts the decimal places in data base to proper number for use with Dec function
            if (Dec == 0.1) { Dec = 1; }
            else if (Dec == 0.01) { Dec = 2; }
            else if (Dec == 0.001) { Dec = 3; }
            else if (Dec == 0.0001) { Dec = 4; }
            else if (Dec == 0.25) { Dec = 2; }
            else if (Dec == 0.5) { Dec = 1; }
            else if (Dec > 1) { Dec = 0; }

            if (RoundSwitch == 1) { Dec = Dec + 5 }

            if (FirstLoop != true) {
                SumStatsJson += ',[';
            }
            SumStatsJson += '"' + FinalRptDate + '",';
            SumStatsJson += '"' + OddSample.split(' ').join('') + '",';
            SumStatsJson += '"' + OddSample.split(' ').join('') + '",';// Twice for charts on client side
            SumStatsJson += '"' + TestTitle + '",';
            SumStatsJson += '"' + Unit + '",';
            SumStatsJson += '"' + NumOfLabs + '",';
            SumStatsJson += '"' + parseFloat(RoundASTM(OddAvg, Dec + 1)).toFixed(Dec) + '",';
            // Add one decemal place to the SD if Decimal places are 0 or 4	
            if (Dec == 0 || Dec == 4 || RoundSwitch == 1) { SumStatsJson += '"' + parseFloat(RoundASTM(OddSD, Dec + 1)).toFixed(Dec) + '",'; }
            else { SumStatsJson += '"' + parseFloat(RoundASTM(OddSD, Dec + 1)).toFixed(Dec + 1) + '",'; }
            SumStatsJson += '"' + parseFloat(RoundASTM(OddCV, Dec + 1)).toFixed(Dec) + '",';
            // Add one decemal place to the SD if Decimal places are 0 or 4	
            SumStatsJson += '"' + parseFloat(RoundASTM(EvenAvg, Dec + 1)).toFixed(Dec) + '",';
            if (Dec == 0 || Dec == 4 || RoundSwitch == 1) { SumStatsJson += '"' + parseFloat(RoundASTM(EvenSD, Dec + 1)).toFixed(Dec) + '",'; }
            else { SumStatsJson += '"' + parseFloat(RoundASTM(EvenSD, Dec + 1)).toFixed(Dec + 1) + '",'; }
            SumStatsJson += '"' + parseFloat(RoundASTM(EvenCV, Dec + 1)).toFixed(Dec) + '",';
            SumStatsJson += '"' + Outliers + '",';
            SumStatsJson += '"' + TestNumber + '"]';
            FirstLoop = false;
            RecordSet.MoveNext;
        };
        SumStatsJson += ']}';
        Response.Write(SumStatsJson);
    }// End display summery of stats.**********************************************************************************************
    else if (NavSwitch == 'Prelim')// Displays preliminary report
    {
        // Begining of Stats function used for Prelim report.
        // Call this function to do all the stats ( Avg SD CV )  
        // Input an array (Data); Outputs an object (Result) with three properties (Result.AVG), (Result.SD), (Result.CV).		
        function Stats(Data) {
            var Result = { Avg: 0, SD: 0, CV: 0 }, t = Data.length;
            for (var m, s = 0, l = t; l--; s += parseFloat(Data[l]));
            for (m = Result.Avg = s / t, l = t, s = 0; l--; s += Math.pow(Data[l] - m, 2));
            Result.SD = Math.sqrt(s / t)
            return Result.CV = 100 * Result.SD / Result.Avg, Result
            // We can do the varience as well if you want.
            //return Result.SD = Math.sqrt(Result.Varience = s / t), Result					
        };
        // End of Stats function

        // These variables are used for prelimary report
        //  This variable will hold all the html code and display the prelimary report page.
        var DisplayPrelim = "<h3>cclr Proficiency Sample Preliminary Report</h3></tr></thead>";
        var LabNum = [];
        var TestName = [];
        var OutLierName = [];
        var OutLiers = [];
        var LabNum = [];
        var OddData = [];
        var EvenData = [];
        // Record set Arrays and associated variables used for Preliminary report.
        var AllDataInSamp = [];  // Used for Preliminary report
        var TestCount = null; // Used to count the number of tests in a program from the  Data table.
        var CurrentOddSample = 0;// Used for heading of Prelim Report

        // Database query used for Pelimary report
        CmdTextStr = 'DECLARE @Program AS Char(10);'
        CmdTextStr += 'SELECT @Program = ?;'
        CmdTextStr = "SELECT * From " + Program + "Data";//Sets the SQL connection strings
        TheComm = Server.CreateObject("ADODB.Command");// Create a command object.
        TheComm.ActiveConnection = TheConn;// Assine the comand object to the connectin object TheConn.
        TheComm.CommandText = CmdTextStr;// Assine the the comand text to the command object TheComm.	
        TheComm.Parameters.Append(TheComm.CreateParameter('@Program', 129, 1, 10, Program));
        RecordSet = TheComm.Execute(adCmdText + adExecuteNoRecords);// Record set from parameterized query table Data to retrieve all data for the selected program.
        TheComm = null;// Release the command object.

        //  This while reads in the data from a specific program in ProgInfo table of pspdata data base into an array of arrays "ProgTest".
        TestCount = RecordSet.Fields.Count;
        g = 0;
        CurrentOddSample = RecordSet.Fields('OddSample').value
        ChoseSample = CurrentOddSample
        EvenSample = CurrentOddSample
        EvenSample++;
        while (!RecordSet.EOF) {
            AllDataInSamp[g] = [];
            AllDataInSamp[g]['LabNum'] = RecordSet.Fields('LabNum').value;
            for (j = 9; j < TestCount; j++) {
                AllDataInSamp[g][RecordSet.Fields(j).name] = RecordSet.Fields(j).value;
            };
            g++
            RecordSet.Move(4);
        }
        RecordSet.close();

        // Boots smarties out to home if data collection is open!!
        var NewDate = new Date()
        if (NewDate < ProgTest[0]['CloseDate']) {
            Response.Redirect('ProgSelect.asp?NavSwitch=SignOut')
        }

        var FirstLoop = true;
        var PrelimJson = "";
        PrelimJson += '{"NavSwitch":"' + NavSwitch + '", "Material":"' + Material + '", "SampleType":"' + SampleType + '", "OddSample":' + CurrentOddSample + ', "data":[[';

        for (j in ProgTest) {
            if (ProgTest[j]['DisplayStatus'].slice(0, 2) != 'NN') // Only add test heading if at least one sample is to be displayed.
            {
                TestName.push(ProgTest[j]['TestTitle'])
                LabNum = [];
                OddData = [];
                EvenData = [];
                for (k in AllDataInSamp) {
                    if (AllDataInSamp[k]['Test' + ProgTest[j]['TestNumber'] + 'Odd'] != '' && AllDataInSamp[k]['Test' + ProgTest[j]['TestNumber'] + 'Eve'] != ''
                        && AllDataInSamp[k]['Test' + ProgTest[j]['TestNumber'] + 'Odd'] != ' ' && AllDataInSamp[k]['Test' + ProgTest[j]['TestNumber'] + 'Eve'] != ' '
                        && AllDataInSamp[k]['Test' + ProgTest[j]['TestNumber'] + 'Odd'] != null && AllDataInSamp[k]['Test' + ProgTest[j]['TestNumber'] + 'Eve'] != null
                        && AllDataInSamp[k]['Test' + ProgTest[j]['TestNumber'] + 'Odd'] != 'BadData' && AllDataInSamp[k]['Test' + ProgTest[j]['TestNumber'] + 'Eve'] != 'BadData'
                        && AllDataInSamp[k]['Test' + ProgTest[j]['TestNumber'] + 'Odd'] != 'undefined' && AllDataInSamp[k]['Test' + ProgTest[j]['TestNumber'] + 'Eve'] != 'undefined') {
                        LabNum.push([j], [AllDataInSamp[k]['LabNum']])
                        OddData.push(AllDataInSamp[k]['Test' + ProgTest[j]['TestNumber'] + 'Odd'])
                        EvenData.push(AllDataInSamp[k]['Test' + ProgTest[j]['TestNumber'] + 'Eve'])
                    }
                }
                var OddStats = (Stats(OddData));
                var EvenStats = (Stats(EvenData));
                // Converts the decimal places in data base to proper number for use with Dec function.
                var Dec = ProgTest[j]['TestDec']
                if (Dec == 0.1) { Dec = 1; }
                else if (Dec == 0.01) { Dec = 2; }
                else if (Dec == 0.001) { Dec = 3; }
                else if (Dec == 0.0001) { Dec = 4; }
                else if (Dec == 0.25) { Dec = 2; }
                else if (Dec == 0.5) { Dec = 1; }
                else if (Dec > 1) { Dec = 0; }
                // Creates copyies of the arrays to put throught the "OutLier removal process"
                var LabNum3SD = LabNum.slice();
                var OddData3SD = OddData.slice();
                var EvenData3SD = EvenData.slice();
                // Calculate the stats to use to Run through the Loop and remove the array elements 3 + standard deviations from mean.
                var OddStats3SD = (Stats(OddData3SD));
                var EvenStats3SD = (Stats(EvenData3SD));
                var Run = 1
                var Loop = 0
                while (Loop < Run) {
                    var OutTemp = '';
                    for (i in OddData3SD) {
                        if (Math.abs(OddData3SD[i]) > 3 * OddStats3SD.SD + OddStats3SD.Avg || Math.abs(EvenData3SD[i]) > 3 * EvenStats3SD.SD + EvenStats3SD.Avg) {
                            OutTemp += ', ' + LabNum3SD[i];
                            LabNum3SD.splice(i, 1)
                            OddData3SD.splice(i, 1)
                            EvenData3SD.splice(i, 1)
                        };
                    };
                    OutTemp = OutTemp.substring(2)
                    OutLiers.push(OutTemp)
                    var OddStats3SD = (Stats(OddData3SD));
                    var EvenStats3SD = (Stats(EvenData3SD));
                    Loop++
                }

                if (FirstLoop != true) {
                    PrelimJson += ',[';
                }
                PrelimJson += '"' + ProgTest[j]['TestTitle'] + '",';
                PrelimJson += '"' + ProgTest[j]['TestUnit'] + '",';
                PrelimJson += '"' + parseFloat(RoundASTM(OddStats.Avg, Dec)).toFixed(Dec) + '",';
                PrelimJson += '"' + parseFloat(RoundASTM(OddStats.SD, Dec + 1)).toFixed(Dec) + '",';
                PrelimJson += '"' + parseFloat(RoundASTM(OddStats.CV, Dec + 1)).toFixed(Dec) + '",';
                PrelimJson += '"' + parseFloat(RoundASTM(EvenStats.Avg, Dec)).toFixed(Dec) + '",';
                PrelimJson += '"' + parseFloat(RoundASTM(EvenStats.SD, Dec + 1)).toFixed(Dec) + '",';
                PrelimJson += '"' + parseFloat(RoundASTM(EvenStats.CV, Dec + 1)).toFixed(Dec) + '"]';
                FirstLoop = false;
            }
        }
        PrelimJson += ']}'
        Response.Write(PrelimJson);
    }
    RecordSet = null;
    Close_Connection();
</script>
