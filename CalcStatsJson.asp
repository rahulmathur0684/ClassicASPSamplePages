<%@  language="JScript" %>
<!--#include file="../ConnectionString/Connection.asp" -->
<script language="JavaScript" runat="server">
    // Record set Arrays and associated variables
    var RecordSet = null;  // Used to store any record set from parameterized query untill the record set is  placed into an array of arrayes.
    var CmdTextStr = null; // The command text string for the command object.
    var g = null; // variable used to loop arrays
    var h = null; // variable used to loop arrays
    var i = null; // variable used to loop arrays
    var j = null; // variable used to loop arrays
    var k = null; // variable used to loop arrays
    var JsonResponse = ''; // this json string will be written to the client once all data is prepared.
    var Now = new Date();
    var Material = null;
    var SampleType = null;
    var FinalRptDate = null;
    var ProgTest = []; // Used for all three reports, Ratings Summmation of Stats and Prelim Report.
    var AllDataInSamp = []; // this will come from the data table.
    var SampleStats = []; // this will come from the Stats table.
    var TestName = [];
    var OutLierName = [];
    var OutLiers = [];
    var LabNum = [];
    var OddData = [];
    var EvenData = [];

    if (Request('Program') == 'GetPrograms') {
        Program = GetPrograms('GetPrograms');
        Result = GetAllDataInSamp(Program);
        GetTests(Program);
        ProduceJson(Program, Result.CurrentOddSample, Result.AllDataInSamp, ProgTest);
    }
    else if (Request('Job') == 'UpLoadStats') {
        var Program = Request('Program');
        GetPrograms(Program);
        Result = GetAllDataInSamp(Program);
        GetTests(Program);
        RawStats = CalcRawStats(Result.AllDataInSamp, ProgTest);
        UploadStats(RawStats);
        ProduceJson(Program, Result.CurrentOddSample, Result.AllDataInSamp, ProgTest);
    }
    else {
        var Program = Request('Program');
        GetPrograms(Program);
        Result = GetAllDataInSamp(Program);
        GetTests(Program);
        ProduceJson(Program, Result.CurrentOddSample, Result.AllDataInSamp, ProgTest);
    }
    // Functions
    function GetPrograms(Program) {
        CmdTextStr = 'SELECT Program, Material, SampleType, CloseDate FROM ProgramInfo Where (DisplayOrder = 1);';
        TheComm = Server.CreateObject("ADODB.Command");// Create a command object.
        TheComm.ActiveConnection = TheConn;// Assine the comand object to the connectin object TheConn.
        TheComm.CommandText = CmdTextStr;// Assine the the comand text to the command object TheComm.	

        // Use parameterized query to stop SQL Injection attacks	
        RecordSet = TheComm.Execute(adCmdText + adExecuteNoRecords);// Record set from parameterized query table "ProgramInfo" to retrieve heading info for the selected program.
        TheComm = null;// Release the command object.
        JsonResponse += '{"Programs":[';
        var CloseData = [];
        g = 0;
        // This while reads in all the programs in ProgramInfo. this will be used for the select box on the client.
        while (!RecordSet.EOF) {
            CloseData[g] = [];
            JsonResponse += '"' + RecordSet.Fields('Program').value + '",';
            CloseDate = new Date(RecordSet.Fields('CloseDate').value);
            if (Now > CloseDate) {
                CloseData[g].Program = RecordSet.Fields('Program').value;
                CloseData[g].CloseDate = RecordSet.Fields('CloseDate').value;
                g++;
            }
            RecordSet.MoveNext;
        }
        RecordSet.close();
        JsonResponse = JsonResponse.substr(0, JsonResponse.length - 1); // remove last troublesome ",".
        JsonResponse += '],';
        // sort in order of CloseDate for the obvious reason. 
        CloseData.sort(function (a, b) { return b.CloseDate - a.CloseDate; });
        var Program = CloseData[0].Program;    
        return Program
    }

    function GetAllDataInSamp(Program) {
        g = 0; // variable used to loop arrays
        CmdTextStr = 'DECLARE @Program AS Char(10);'
        CmdTextStr += 'SELECT @Program = ?;'
        CmdTextStr = 'SELECT * From ' + Program + 'Data';//Sets the SQL connection strings
        TheComm = Server.CreateObject("ADODB.Command");// Create a command object.

        TheComm.ActiveConnection = TheConn;// Assine the comand object to the connectin object TheConn.
        TheComm.CommandText = CmdTextStr;// Assine the the comand text to the command object TheComm.	
        TheComm.Parameters.Append(TheComm.CreateParameter('@Program', 129, 1, 10, Program));
        RecordSet = TheComm.Execute(adCmdText + adExecuteNoRecords);// Record set from parameterized query table Data to retrieve all data for the selected program.
        TheComm = null;// Release the command object.
        TestCount = RecordSet.Fields.Count;
        CurrentOddSample = RecordSet.Fields('OddSample').value
        EvenSample = CurrentOddSample
        EvenSample++;
        while (!RecordSet.EOF) {
            AllDataInSamp[g] = [];
            if (RecordSet.Fields('LabNum').value != 0) {
                AllDataInSamp[g]['LabNum'] = RecordSet.Fields('LabNum').value;
                for (j = 9; j < TestCount; j++) {
                    AllDataInSamp[g][RecordSet.Fields(j).name] = RecordSet.Fields(j).value;
                }
                g++;
            }
            RecordSet.Move(1);
        }
        RecordSet.close();
        var Result = { AllDataInSamp: AllDataInSamp, CurrentOddSample: CurrentOddSample }
        return Result;
    }

    // Get records
    function GetTests(Program) {
        // Now we get all the info for the chosen program
        CmdTextStr = 'DECLARE @Program AS Char(10);'
        CmdTextStr += 'SELECT @Program = ?;'
        CmdTextStr += 'SELECT Program, Material, SampleType, CloseDate, TestNumber, TestTitle, DisplayOrder, DisplayStatus, TestUnit, TestDec FROM ProgramInfo Where (Program = @Program)';
        TheComm = Server.CreateObject("ADODB.Command");// Create a command object.
        TheComm.ActiveConnection = TheConn;// Assine the comand object to the connectin object TheConn.
        TheComm.CommandText = CmdTextStr;// Assine the the comand text to the command object TheComm.	

        // Use parameterized query to stop SQL Injection attacks

        TheComm.Parameters.Append(TheComm.CreateParameter('@Program', 129, 1, 10, Program)); // Create the parameter query based on the variable Program. Can be Concrete, Portland Cement ect. 
        RecordSet = TheComm.Execute(adCmdText + adExecuteNoRecords);// Record set from parameterized query table "ProgramInfo" to retrieve heading info for the selected program.
        TheComm = null;// Release the command object.

        // This while reads in the data from a specific program in ProgramInfo table of pspdata data base into an array of arrays "ProgTest".
        g = 0;
        while (!RecordSet.EOF) {
            if (RecordSet.Fields('DisplayStatus').value.substr(2, 1) != 'N') {
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
                ProgTest[g]['Stats'] = [];
                g++
            }
            RecordSet.MoveNext;
        }
        RecordSet.close();
        Material = ProgTest[0]['Material'];
        SampleType = ProgTest[0]['SampleType'];
    }

    function CalcRawStats(AllDataInSamp, ProgTest) {
        var RawStats = [];
        i: for (var j = 0; j < ProgTest.length; j++) {
            var LabCount = 0;
            var DataOdd = [];
            var DataEven = [];
            var TestNumber = ProgTest[j]['TestNumber'];
            RawStats[TestNumber] = [];
            for (var k = 0; k < AllDataInSamp.length; k++) {
                if (AllDataInSamp[k]['LabNum'] == 0) { break i; }
                if (AllDataInSamp[k]['Test' + TestNumber + 'Odd'] != '' && AllDataInSamp[k]['Test' + TestNumber + 'Eve'] != ''
                    && AllDataInSamp[k]['Test' + TestNumber + 'Odd'] != ' ' && AllDataInSamp[k]['Test' + TestNumber + 'Eve'] != ' '
                    && AllDataInSamp[k]['Test' + TestNumber + 'Odd'] != null && AllDataInSamp[k]['Test' + TestNumber + 'Eve'] != null
                    && AllDataInSamp[k]['Test' + TestNumber + 'Odd'] != 'BadData' && AllDataInSamp[k]['Test' + TestNumber + 'Eve'] != 'BadData'
                    && AllDataInSamp[k]['Test' + TestNumber + 'Odd'] != 'undefined' && AllDataInSamp[k]['Test' + TestNumber + 'Eve'] != 'undefined') {
                    DataOdd.push(AllDataInSamp[k]['Test' + ProgTest[j]['TestNumber'] + 'Odd']);
                    DataEven.push(AllDataInSamp[k]['Test' + ProgTest[j]['TestNumber'] + 'Eve']);
                    LabCount++;
                }
            }
            RawStats[TestNumber]['NumOfLabs'] = LabCount;
            RawStats[TestNumber]['OddStats'] = Stats(DataOdd);
            RawStats[TestNumber]['EvenStats'] = Stats(DataEven);
        }
        return RawStats;
    }

    // Begining of Stats function.
    // Call this function to do all the stats (Avg, SD, CV, Varience)  
    // Input an array (Data) 
    // Outputs an object (Result) with four properties Result.AVG, Result.SD corrected (n-1), Result.CV and Result.Var.
    function Stats(Data) {
        var Result = { Avg: 0, SD: 0, CV: 0, Var: 0 }, n = Data.length;
        for (var avg, sqdf = 0, l = n; l--; sqdf += parseFloat(Data[l]));
        for (avg = Result.Avg = sqdf / n, l = n, sqdf = 0; l--; sqdf += Math.pow(Data[l] - avg, 2));
        Result.SD = Math.sqrt(sqdf / (n - 1)); Result.CV = 100 * Result.SD / Result.Avg; //Result.Var = sqdf / n; cclr doesn't use variance.
        return Result;
    }	// End of Stats function

    function UploadStats(RawStats) {
        var Program = Request('Program');
        var OddSample = Request('OddSample');
        var Display = Request('Display');
        var TestNumber = Request('TestNumber');
        var FinalRptDate = Request('FinalRptDate');

        DeleteStats(Program, OddSample, TestNumber);
        for (g = 0; g < 2; g++) {
            if (g == 0) {
                var NumOfLabs = RawStats[TestNumber]['NumOfLabs'];
                var OddAvg = parseFloat(RawStats[TestNumber]['OddStats'].Avg).toFixed(5);
                var OddSD = parseFloat(RawStats[TestNumber]['OddStats'].SD).toFixed(5);
                var EvenAvg = parseFloat(RawStats[TestNumber]['EvenStats'].Avg).toFixed(5);
                var EvenSD = parseFloat(RawStats[TestNumber]['EvenStats'].SD).toFixed(5);
                var Outliers = 'All Data';
            }
            else {
                var Outliers = Request('Outliers') + '';
                Outliers = Outliers.split(',')
                var NumOfOutliers = 0;
                if (Request('Outliers') != '') { NumOfOutliers = Outliers.length }
                var NumOfLabs = RawStats[TestNumber]['NumOfLabs'] - NumOfOutliers;
                NumOfLabs = NumOfLabs + '  @ ' + Request('Run') + '';
                var OddAvg = Request('OddAvg');
                var OddSD = Request('OddSD');
                var EvenAvg = Request('EvenAvg');
                var EvenSD = Request('EvenSD');
            }
            CmdTextStr = 'DECLARE @OddSample AS Char(3);'
            CmdTextStr += 'DECLARE @Display AS Char(4);'
            CmdTextStr += 'DECLARE @TestNumber AS Char(5);'
            CmdTextStr += 'DECLARE @NumOfLabs AS Char(10);'
            CmdTextStr += 'DECLARE @OddSD AS Char(15);'
            CmdTextStr += 'DECLARE @EvenSD AS Char(15);'
            CmdTextStr += 'DECLARE @OddAvg AS Char(15);'
            CmdTextStr += 'DECLARE @EvenAvg AS Char(15);'
            CmdTextStr += 'DECLARE @Outliers AS Char(1000);'
            CmdTextStr += 'DECLARE @FinalRptDate AS Char(30)';

            CmdTextStr += 'SELECT @OddSample = ?;'
            CmdTextStr += 'SELECT @Display = ?;'
            CmdTextStr += 'SELECT @TestNumber = ?;'
            CmdTextStr += 'SELECT @NumOfLabs = ?;'
            CmdTextStr += 'SELECT @OddSD = ?;'
            CmdTextStr += 'SELECT @EvenSD = ?;'
            CmdTextStr += 'SELECT @OddAvg = ?;'
            CmdTextStr += 'SELECT @EvenAvg = ?;'
            CmdTextStr += 'SELECT @Outliers = ?;'
            CmdTextStr += 'SELECT @FinalRptDate = ?;'

            CmdTextStr += 'Insert Into ' + Program + 'Stats (OddSample, Display, TestNumber, NumOfLabs, OddSD, EvenSD, OddAvg, EvenAvg, Outliers, FinalRptDate)';
            CmdTextStr += 'Values (@OddSample, @Display, @TestNumber, @NumOfLabs, @OddSD, @EvenSD, @OddAvg, @EvenAvg, @Outliers, @FinalRptDate)';

            TheComm = Server.CreateObject("ADODB.Command");// Create a command object.
            TheComm.ActiveConnection = TheConn;// Assine the comand object to the connectin object TheConn.
            TheComm.CommandText = CmdTextStr;// Assine the the comand text to the command object TheComm.	
            // Use parameterized query to stop SQL Injection attacks
            TheComm.Parameters.Append(TheComm.CreateParameter('@OddSample', 129, 1, 3, OddSample));
            TheComm.Parameters.Append(TheComm.CreateParameter('@Display', 129, 1, 4, Display));
            TheComm.Parameters.Append(TheComm.CreateParameter('@TestNumber', 129, 1, 5, TestNumber));
            TheComm.Parameters.Append(TheComm.CreateParameter('@NumOfLabs', 129, 1, 10, NumOfLabs));
            TheComm.Parameters.Append(TheComm.CreateParameter('@OddSD', 129, 1, 12, OddSD));
            TheComm.Parameters.Append(TheComm.CreateParameter('@EvenSD', 129, 1, 12, EvenSD));
            TheComm.Parameters.Append(TheComm.CreateParameter('@OddAvg', 129, 1, 12, OddAvg));
            TheComm.Parameters.Append(TheComm.CreateParameter('@EvenAvg', 129, 1, 12, EvenAvg));
            TheComm.Parameters.Append(TheComm.CreateParameter('@Outliers', 129, 1, 1000, Outliers));
            TheComm.Parameters.Append(TheComm.CreateParameter('@FinalRptDate', 129, 1, 30, FinalRptDate));
            RecordSet = TheComm.Execute(adCmdText + adExecuteNoRecords);// Record set from parameterized query table "ProgramInfo" to retrieve heading info for the selected program.
            TheComm = null;// Release the command object.
        }
    }

    function DeleteStats(Program, OddSample, TestNumber) {
        CmdTextStr = 'DECLARE @OddSample AS Char(3);'
        CmdTextStr += 'DECLARE @TestNumber AS Char(4);'
        CmdTextStr += 'SELECT @OddSample = ?;'
        CmdTextStr += 'SELECT @TestNumber = ?;'
        CmdTextStr += 'DELETE FROM ' + Program + 'Stats Where OddSample = @OddSample AND TestNumber = @TestNumber';
        TheComm = Server.CreateObject("ADODB.Command");// Create a command object.
        TheComm.ActiveConnection = TheConn;// Assine the comand object to the connectin object TheConn.
        TheComm.CommandText = CmdTextStr;// Assine the the comand text to the command object TheComm.	
        // Use parameterized query to stop SQL Injection attacks
        TheComm.Parameters.Append(TheComm.CreateParameter('@OddSample', 129, 1, 3, OddSample));
        TheComm.Parameters.Append(TheComm.CreateParameter('@TestNumber', 129, 1, 4, TestNumber));
        RecordSet = TheComm.Execute(adCmdText + adExecuteNoRecords);// Record set from parameterized query table "ProgramInfo" to retrieve heading info for the selected program.
        TheComm = null;// Release the command object.
    }

    function ProduceJson(Program, CurrentOddSample, AllDataInSamp, ProgTest) {
        // now we get all the Stats (if they have been preveasly uploaded for the chosen program.
        CmdTextStr = 'DECLARE @OddSample AS Char(4);'
        CmdTextStr += 'SELECT @OddSample = ?;'
        CmdTextStr += 'SELECT Display, TestNumber From ' + Program + 'Stats Where (OddSample = @OddSample);'; //Sets the SQL connection stringsTheComm = Server.CreateObject("ADODB.Command");// Create a command object.
        TheComm = Server.CreateObject("ADODB.Command");// Create a command object.
        TheComm.ActiveConnection = TheConn;// Assine the comand object to the connectin object TheConn.
        TheComm.CommandText = CmdTextStr;// Assine the the comand text to the command object TheComm.	
        // Use parameterized query to stop SQL Injection attacks

        TheComm.Parameters.Append(TheComm.CreateParameter('@OddSample', 202, 1, 4, CurrentOddSample - 2)); // Create the parameter query based on the variable Program. Can be Concrete, Portland Cement ect. 
        RecordSet = TheComm.Execute(adCmdText + adExecuteNoRecords);// Record set from parameterized query table "ProgramInfo" to retrieve heading info for the selected program.
        TheComm = null;// Release the command object.

        // sort in order of DispayOrder for the obvious reason. 
        ProgTest.sort(function (a, b) { return a.DisplayOrder - b.DisplayOrder; });
        g = 0;
        while (!RecordSet.EOF) {
            h: for (var g = 0; g < ProgTest.length; g++)if (ProgTest[g]['TestNumber'] == RecordSet.Fields('TestNumber').value.split(' ').join('')) {
                ProgTest[g]['Display'] = RecordSet.Fields('Display').value;
                ProgTest[g]['Stats']['OddSD'] = 'NotUploaded';
                ProgTest[g]['Stats']['EveSD'] = 'NotUploaded';
                ProgTest[g]['Stats']['OddAvg'] = 'NotUploaded';
                ProgTest[g]['Stats']['EveAvg'] = 'NotUploaded';
                break h;
            }
            RecordSet.Move(1);
        }
        RecordSet.close();

        // now we get all the Stats (if they have been preveasly uploaded for the chosen program.
        CmdTextStr = 'DECLARE @OddSample AS Char(4);'
        CmdTextStr += 'SELECT @OddSample = ?;'
        CmdTextStr += 'SELECT * From ' + Program + 'Stats Where (OddSample = @OddSample);'; //Sets the SQL connection stringsTheComm = Server.CreateObject("ADODB.Command");// Create a command object.
        TheComm = Server.CreateObject("ADODB.Command");// Create a command object.
        TheComm.ActiveConnection = TheConn;// Assine the comand object to the connectin object TheConn.
        TheComm.CommandText = CmdTextStr;// Assine the the comand text to the command object TheComm.	
        // Use parameterized query to stop SQL Injection attacks

        TheComm.Parameters.Append(TheComm.CreateParameter('@OddSample', 202, 1, 4, CurrentOddSample)); // Create the parameter query based on the variable Program. Can be Concrete, Portland Cement ect. 
        RecordSet = TheComm.Execute(adCmdText + adExecuteNoRecords);// Record set from parameterized query table "ProgramInfo" to retrieve heading info for the selected program.
        TheComm = null;// Release the command object.

        while (!RecordSet.EOF) {
            if (RecordSet.Fields('FinalRptDate').value != null) {
                FinalRptDate = RecordSet.Fields('FinalRptDate').value;
            }

            if (RecordSet.Fields('Outliers').value == null) {
                var Outliers2 = '';
            }
            else {
                var Outliers2 = RecordSet.Fields('Outliers').value.split(' ').join('');
            }
            h: for (var g = 0; g < ProgTest.length; g++)if (ProgTest[g]['TestNumber'] == RecordSet.Fields('TestNumber').value.split(' ').join('')) {
                if (ProgTest[g]['Display'] != RecordSet.Fields('Display').value) {
                    ProgTest[g]['Display'] = RecordSet.Fields('Display').value;
                }
                ProgTest[g]['Stats']['OddSD'] = RecordSet.Fields('OddSD').value;
                ProgTest[g]['Stats']['EveSD'] = RecordSet.Fields('EvenSD').value;
                ProgTest[g]['Stats']['OddAvg'] = RecordSet.Fields('OddAvg').value;
                ProgTest[g]['Stats']['EveAvg'] = RecordSet.Fields('EvenAvg').value;
                if (Outliers2 == '') {
                    ProgTest[g]['Stats']['NumOfOutliers'] = 0;
                }
                else {
                    ProgTest[g]['Stats']['NumOfOutliers'] = RecordSet.Fields('Outliers').value.split(',').length;
                }
                break h;
            }
            RecordSet.Move(1);
        }
        RecordSet.close();

        JsonResponse += '"Program":"' + Program + '","Material":"' + Material + '", "SampleType":"' + SampleType + '", "OddSample":"' + CurrentOddSample + '", "FinalRptDate":"' + FinalRptDate + '",';
        JsonResponse += '"Tests":[';

        i: for (var j = 0; j < ProgTest.length; j++) {
            JsonResponse += '["' + ProgTest[j]['TestTitle'] + '",';
            JsonResponse += '"' + ProgTest[j]['TestNumber'] + '",';
            JsonResponse += '"' + ProgTest[j]['TestUnit'] + '",';
            JsonResponse += '"' + ProgTest[j]['TestDec'] + '",';
            JsonResponse += '"' + ProgTest[j]['Display'] + '",';

            JsonResponse += '{"Stats":[[';
            JsonResponse += '"' + ProgTest[j]['Stats']['OddAvg'] + '",';
            JsonResponse += '"' + ProgTest[j]['Stats']['OddSD'] + '",';
            JsonResponse += '"' + ProgTest[j]['Stats']['EveAvg'] + '",';
            JsonResponse += '"' + ProgTest[j]['Stats']['EveSD'] + '",';
            JsonResponse += '"' + ProgTest[j]['Stats']['NumOfOutliers'] + '"';
            JsonResponse += ']]},'

            JsonResponse += '{"Data":[';

            var LabCount = 0;
            var TestNumRate
            for (var k = 0; k < AllDataInSamp.length; k++) {
                if (AllDataInSamp[k]['LabNum'] == 0) { break i; }
                if (AllDataInSamp[k]['Test' + ProgTest[j]['TestNumber'] + 'Odd'] != '' && AllDataInSamp[k]['Test' + ProgTest[j]['TestNumber'] + 'Eve'] != ''
                    && AllDataInSamp[k]['Test' + ProgTest[j]['TestNumber'] + 'Odd'] != ' ' && AllDataInSamp[k]['Test' + ProgTest[j]['TestNumber'] + 'Eve'] != ' '
                    && AllDataInSamp[k]['Test' + ProgTest[j]['TestNumber'] + 'Odd'] != null && AllDataInSamp[k]['Test' + ProgTest[j]['TestNumber'] + 'Eve'] != null
                    && AllDataInSamp[k]['Test' + ProgTest[j]['TestNumber'] + 'Odd'] != 'BadData' && AllDataInSamp[k]['Test' + ProgTest[j]['TestNumber'] + 'Eve'] != 'BadData'
                    && AllDataInSamp[k]['Test' + ProgTest[j]['TestNumber'] + 'Odd'] != 'undefined' && AllDataInSamp[k]['Test' + ProgTest[j]['TestNumber'] + 'Eve'] != 'undefined') {
                    JsonResponse += '["' + AllDataInSamp[k]['LabNum'] + '",';
                    JsonResponse += '"' + AllDataInSamp[k]['Test' + ProgTest[j]['TestNumber'] + 'Odd'] + '",';
                    JsonResponse += '"' + AllDataInSamp[k]['Test' + ProgTest[j]['TestNumber'] + 'Eve'] + '"],';
                    LabCount++;
                }
            }
            if (LabCount == 0);
            {
                JsonResponse += '["' + 0 + '",';
                JsonResponse += '"' + 1 + '",';
                JsonResponse += '"' + 1 + '"],';
            }
            JsonResponse = JsonResponse.substr(0, JsonResponse.length - 1);
            JsonResponse += ']}],';
        }
        JsonResponse = JsonResponse.substr(0, JsonResponse.length - 1);
        JsonResponse += ']}';
        Response.Write(JsonResponse);
        RecordSet = null;
        Close_Connection();
    }
// End Functions
</script>