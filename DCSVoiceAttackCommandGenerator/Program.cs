// See https://aka.ms/new-console-template for more information

using System.Data;
using System.Text.RegularExpressions;
using HtmlAgilityPack;

///DCS VoiceAttack Command Generator
///DCSVACMDGEN
///Drop your DCS controls html onto DCSCMDGEN.exe and a .vap will be created that represents those commands.
///
///Action Process:
///User makes keybinds in dcs
///User clicks "Make HTML" in DCS controls menu
///User drags the html file onto the DCSVACMDGEN.exe
///A properly formated VA profile is created in the same directory as the exe
///The user imports the generated VAP into their VAP, or can use the result directly.
///
///Code Process:
///The html file is read by the DCSVACMDGEN.exe and formatted into a data table (https://stackoverflow.com/questions/18090626/import-data-from-html-table-to-datatable-in-c-sharp)
///	Generate a GUID for the <ID> (https://learn.microsoft.com/en-us/dotnet/api/system.guid.newguid?view=net-8.0)
///	If there is a keybind associated, then create the Press keybind using keycodes at a rate of some seconds
///	(Maybe) make a gui (or user input in the cmd line) to ask for the command duration. Maybe not. Likely not..
///	Scrub and eliminate the file of all '&'.
///Add the profile prefix and postfix
///Export the result file as a .vap and named as the current date, or something, becuase you can't determine what module and if there is already a file with the same name.
///

// Thank you Voice Attack discord! https://discord.gg/ab42CqD

/* TODO
 * Also be aware of commands that have more than one set of keybinds. A ; shows inbwtween them in the dcs html
 * Also, watch out for commands that have double quotes in them, like this one |1,A/A refueling - "Ready for precontact" radio call|
 * in VA there are 2 ways commands can be messed up. A shortened Spoken Command, and the word "undefined" in the Actions.
 * Watch out for any keybinds that start with =. this will likely cause excel/google sheets to error out and mess up that line for the csv
 * Run Test Cases for things with no binds, 1 bind, 3 binds plus
 * Warn the user what happens if they try to import more than one bind per action
 * figgure out what happens to < and > when put through the program
 * Make a readme
 * Celebrate!
 * Test, test, test.
 */

/* Reach Goals
 * Maybe make some intersting things to determine profile names, key duration, etc
 * Make a logo
 * Based if the file is a html or tsv, do the different processing
 */

/* Bugs/Issues
 * Commands with more than one set of keybinds are not recognized properly. Solution: Warn the user
 */


//https://stackoverflow.com/questions/6261697/what-is-the-best-way-of-converting-a-html-table-to-datatable
var htmlImportFile = args[0]; //this is the name of the file that the user drops onto the exe


/*
if (Path.GetExtension(htmlImportFile) == ".html")
{
    Console.WriteLine("This is an html file.");
    // do html file stuff
}

if (Path.GetExtension(htmlImportFile) == ".tsv")
{
    Console.WriteLine("This is a tsv file.");
    // do tsv file stuff
}
*/


Console.WriteLine("Thank you for using Bailey's DCS Voice Attack Command Generator\nPlease wait...");
var doc = new HtmlDocument();
doc.Load(htmlImportFile);

var nodes = doc.DocumentNode.SelectNodes("//table//tr");
var table = new DataTable("MyTable");

var headers = nodes[0]
    .Elements("th")
    .Select(th => th.InnerText.Trim());

foreach (var header in headers)
{
    table.Columns.Add(header);
}

var rows = nodes.Skip(1).Select(tr => tr
    .Elements("td")
    .Select(td => td.InnerText.Trim())
    .ToArray());
int counter = 0;
foreach (var row in rows)
{
    table.Rows.Add(row);
    counter++;
}

DateTime thisDate = DateTime.Now; // the version of the vap export will be the current dateTime
string filename = Path.GetFileNameWithoutExtension(htmlImportFile); // a part of the export file name and profile name
var vaProfileName = filename + " v" + thisDate.ToString("yyMMddHHmmss"); // v for version

// The prefix for the vap XML is similar for all profiles
var vaPrefix = "<?xml version=\"1.0\"?>" +
    "\r\n<Profile xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">" +
    "\r\n  <HasMB>false</HasMB>" +
    "\r\n  <Id>" + Guid.NewGuid() + "</Id>" + // randomly generated Guid 
    "\r\n  <Name>DCSVACMDGEN " + vaProfileName + "</Name>" + // made from above
    "\r\n  <Commands>";

// The suffix for the vap is similar in all vaps
var vaSuffix = "  </Commands>" +
    "\r\n  <OverrideGlobal>false</OverrideGlobal>" +
    "\r\n  <GlobalHotkeyIndex>0</GlobalHotkeyIndex>" +
    "\r\n  <GlobalHotkeyEnabled>false</GlobalHotkeyEnabled>" +
    "\r\n  <GlobalHotkeyValue>0</GlobalHotkeyValue>" +
    "\r\n  <GlobalHotkeyShift>0</GlobalHotkeyShift>" +
    "\r\n  <GlobalHotkeyAlt>0</GlobalHotkeyAlt>" +
    "\r\n  <GlobalHotkeyCtrl>0</GlobalHotkeyCtrl>" +
    "\r\n  <GlobalHotkeyWin>0</GlobalHotkeyWin>" +
    "\r\n  <GlobalHotkeyPassThru>false</GlobalHotkeyPassThru>" +
    "\r\n  <OverrideMouse>false</OverrideMouse>" +
    "\r\n  <MouseIndex>0</MouseIndex>" +
    "\r\n  <OverrideStop>false</OverrideStop>" +
    "\r\n  <StopCommandHotkeyEnabled>false</StopCommandHotkeyEnabled>" +
    "\r\n  <StopCommandHotkeyValue>0</StopCommandHotkeyValue>" +
    "\r\n  <StopCommandHotkeyShift>0</StopCommandHotkeyShift>" +
    "\r\n  <StopCommandHotkeyAlt>0</StopCommandHotkeyAlt>" +
    "\r\n  <StopCommandHotkeyCtrl>0</StopCommandHotkeyCtrl>" +
    "\r\n  <StopCommandHotkeyWin>0</StopCommandHotkeyWin>" +
    "\r\n  <StopCommandHotkeyPassThru>false</StopCommandHotkeyPassThru>" +
    "\r\n  <DisableShortcuts>false</DisableShortcuts>" +
    "\r\n  <UseOverrideListening>false</UseOverrideListening>" +
    "\r\n  <OverrideJoystickGlobal>false</OverrideJoystickGlobal>" +
    "\r\n  <GlobalJoystickIndex>0</GlobalJoystickIndex>" +
    "\r\n  <GlobalJoystickButton>0</GlobalJoystickButton>" +
    "\r\n  <GlobalJoystickNumber>0</GlobalJoystickNumber>" +
    "\r\n  <GlobalJoystickButton2>0</GlobalJoystickButton2>" +
    "\r\n  <GlobalJoystickNumber2>0</GlobalJoystickNumber2>" +
    "\r\n  <ReferencedProfile xsi:nil=\"true\" />" +
    "\r\n  <ExportVAVersion>1.11</ExportVAVersion>" +
    "\r\n  <ExportOSVersionMajor>10</ExportOSVersionMajor>" +
    "\r\n  <ExportOSVersionMinor>0</ExportOSVersionMinor>" +
    "\r\n  <OverrideConfidence>false</OverrideConfidence>" +
    "\r\n  <Confidence>0</Confidence>" +
    "\r\n  <CatchAllEnabled>false</CatchAllEnabled>" +
    "\r\n  <CatchAllId xsi:nil=\"true\" />" +
    "\r\n  <InitializeCommandEnabled>false</InitializeCommandEnabled>" +
    "\r\n  <InitializeCommandId xsi:nil=\"true\" />" +
    "\r\n  <UseProcessOverride>false</UseProcessOverride>" +
    "\r\n  <ProcessOverrideAciveWindow>true</ProcessOverrideAciveWindow>" +
    "\r\n  <DictationCommandEnabled>false</DictationCommandEnabled>" +
    "\r\n  <DictationCommandId xsi:nil=\"true\" />" +
    "\r\n  <EnableProfileSwitch>false</EnableProfileSwitch>" +
    "\r\n  <GroupCategory>false</GroupCategory>" +
    "\r\n  <LastEditedCommand xsi:nil=\"true\" />" +
    "\r\n  <IS>0</IS>" +
    "\r\n  <IO>0</IO>" +
    "\r\n  <IP>0</IP>" +
    "\r\n  <BE>0</BE>" +
    "\r\n  <UnloadCommandEnabled>false</UnloadCommandEnabled>" +
    "\r\n  <UnloadCommandId xsi:nil=\"true\" />" +
    "\r\n  <BlockExternal>false</BlockExternal>" +
    "\r\n  <AuthorID xsi:nil=\"true\" />" +
    "\r\n  <ProductID xsi:nil=\"true\" />" +
    "\r\n  <CR>0</CR>" +
    "\r\n  <InternalID xsi:nil=\"true\" />" +
    "\r\n  <PR>0</PR>" +
    "\r\n  <CO>0</CO>" +
    "\r\n  <OP>0</OP>" +
    "\r\n  <CV>0</CV>" +
    "\r\n  <PD>0</PD>" +
    "\r\n  <PE>0</PE>" +
    "\r\n  <ExecOnRecognizedEnabled>false</ExecOnRecognizedEnabled>" +
    "\r\n  <ExecOnRecognizedId xsi:nil=\"true\" />" +
    "\r\n  <ExecOnRecognizedRejected>false</ExecOnRecognizedRejected>" +
    "\r\n  <ExcludeGlobalProfiles>false</ExcludeGlobalProfiles>" +
    "\r\n  <DisableAdvancedTTS>false</DisableAdvancedTTS>" +
    "\r\n  <RPR>0</RPR>" +
    "\r\n  <Deleted>false</Deleted>" +
    "\r\n</Profile>";

// used to give the user an idea of how many commands should be present when they open the profile in VA
int commandCreationCounter = 0; 

// gets the location of the exe as a part of the export path
string exeLocation = AppDomain.CurrentDomain.BaseDirectory;

// make the ful path. works where ever the exe might end up
string fullPath = exeLocation + "DCSVACMDGEN " + vaProfileName + ".vap";

// starting writing to file
using (StreamWriter writer = new StreamWriter(fullPath))
{
    writer.WriteLine(vaPrefix); // use the prefix that was defined earlier
    // this will evaluate one line at a time
    // the counter is populdated from when we made the rows from the html file
    for (int i = 0; i < counter; i++)
    {
        // the first part are the keybind combos, eg, "LShift - G". The second part are the Commands, eg, "Gear Up"
        //Console.WriteLine(table.Rows[i]["Combos"] + "|" + table.Rows[i]["Command"]); // debug
        string currentCommand = (string)table.Rows[i]["Command"];
        string currentKeycombo = (string)table.Rows[i]["Combos"];

        // you can use https://dotnetfiddle.net/ to test bits of code
        // you can use https://regex101.com/ to tst regex

        // the code below is for tab seperated values
        // string[] partsOfLine = Regex.Split(line, @"(.*)\t(.*)"); //split the useful parts. Format is (Keycombo),(Action)

        //if (partsOfLine.Length > 2)//this will trigger only if the above regex was good, I guess
        //{
        // partsOfLine[0] is nothing
        // currentKeycombo is the Keycombo
        // currentCommand is the Action

        // Universal Header
        writer.WriteLine("    <Command>" +
        "\r\n      <Referrer xsi:nil=\"true\" />" +
        "\r\n      <ExecType>3</ExecType>" +
        "\r\n      <Confidence>0</Confidence>" +
        "\r\n      <PrefixActionCount>0</PrefixActionCount>" +
        "\r\n      <IsDynamicallyCreated>false</IsDynamicallyCreated>" +
        "\r\n      <TargetProcessSet>false</TargetProcessSet>" +
        "\r\n      <TargetProcessType>0</TargetProcessType>" +
        "\r\n      <TargetProcessLevel>0</TargetProcessLevel>" +
        "\r\n      <CompareType>0</CompareType>" +
        "\r\n      <ExecFromWildcard>false</ExecFromWildcard>" +
        "\r\n      <IsSubCommand>false</IsSubCommand>" +
        "\r\n      <IsOverride>false</IsOverride>" +
        "\r\n      <BaseId>" + Guid.NewGuid() + "</BaseId>" + // new Guid
        "\r\n      <OriginId>00000000-0000-0000-0000-000000000000</OriginId>" +
        "\r\n      <SessionEnabled>true</SessionEnabled>" +
        "\r\n      <DoubleTapInvoked>false</DoubleTapInvoked>" +
        "\r\n      <SingleTapDelayedInvoked>false</SingleTapDelayedInvoked>" +
        "\r\n      <LongTapInvoked>false</LongTapInvoked>" +
        "\r\n      <ShortTapDelayedInvoked>false</ShortTapDelayedInvoked>" +
        "\r\n      <SleepFlag>0</SleepFlag>" +
        "\r\n      <Id>" + Guid.NewGuid() + "</Id>"); // new Guid

        // process out all the weird characters that mess things up
        // because we are using html import, you can keep the commas in
        // the contains for the double quotes is in there because when using csv they are the excape for fields with commas
        // TODO elimate unneeded code here
        if (currentCommand.Contains("&") || currentCommand.Contains("\"") || currentCommand.Contains(","))
        {
            string partsOfLine2 = currentCommand.Replace("&", "and ").Replace("and  ", "and ").Replace("\"", "").Replace(",", "");
            // the order of the replace is important
            // the replacement of double quotes is a VA requirement
            writer.WriteLine("      <CommandString>" + partsOfLine2 + "</CommandString>"); // the name of the command
        }
        else // it does not have odd characters that mess things up
        {
            writer.WriteLine("      <CommandString>" + currentCommand + "</CommandString>"); // the name of the command
        }

        // The differences for if the command has an action or not start here

        if (currentKeycombo == "")// if there are no keybinds associated
        {
            writer.WriteLine("<ActionSequence />");
        }
        else // there are keybinds that go with the command
        {
            writer.WriteLine("\r\n      <ActionSequence>" +
        "\r\n        <CommandAction>" +
        "\r\n          <_caption>Placeholder01</_caption>" + // It seems like these captions do not matter
                                                             // They would tylically say something like:
                                                             // Press F1 key and hold for 0.115 seconds and release

        "\r\n          <PairingSet>false</PairingSet>" +
        "\r\n          <PairingSetElse>false</PairingSetElse>" +
        "\r\n          <Ordinal>0</Ordinal>" +
        "\r\n          <ConditionMet xsi:nil=\"true\" />" +
        "\r\n          <IndentLevel>0</IndentLevel>" +
        "\r\n          <ConditionSkip>false</ConditionSkip>" +
        "\r\n          <IsSuffixAction>false</IsSuffixAction>" +
        "\r\n          <DecimalTransient1>0</DecimalTransient1>" +
        "\r\n          <Caption>Placeholder02</Caption>" + // It seems like these captions do not matter
        "\r\n          <Id>" + Guid.NewGuid() + "</Id>" + // new Guid
        "\r\n          <ActionType>PressKey</ActionType>" +
        "\r\n          <Duration>0.216</Duration>" + // this can be changed
        "\r\n          <Delay>0</Delay>" +
        "\r\n          <KeyCodes>");

            //"\r\n            <unsignedShort>71</unsignedShort>" + // this is where we put the keys that this applies to
            // When processing the key combo, get the unsigned shorts and then stack them here as a list

            string partsOfLine1;
            // if the csv had a keybind that had a comma, it will put the keys in double quotes. get rid of them
            // this should not be necessary when using html import
            // TODO evaluate if this is required when using HTML import
            if (currentKeycombo.EndsWith("\"") && currentKeycombo.StartsWith("\""))
            {
                partsOfLine1 = currentKeycombo.Replace("\"", ""); // replace the double quotes with nothing
            }
            else
            {
                partsOfLine1 = currentKeycombo; // else do nothing
            }

            // TODO around this area you can check to see if more than 1 key combo is listed, example:
            // LShift - G;RCtrl - A
            // If the above is the case, then use the first of the keybinds
            // There is a case where the char ; is used for a keybind. That looks like this:
            // Lshift - ;
            // and combined they look like this:
            // Lshift - ;;RCtrl - A
            // or
            // RCtrl - A;LShift - ;
            // or
            // RCtrl - A;;
            // or
            // ;;RCtrl - A
            // jeez. Ok, maybe just say that multiple keyboard binds are ow supported at this time...
            // For the ones that are suject to this, you can bind them manually in VA.

            // break aart the elements of the keybind
            // the - char is used as a separator in DCS and always has a space before and after
            string[] keycombos = Regex.Split(partsOfLine1, @"( - )");

            foreach (string key in keycombos)
            {
                if (key.Equals(" - "))
                {
                    //do nothing 
                }
                else
                {
                    writer.WriteLine("            <unsignedShort>" + GetKeycode(key) + "</unsignedShort>");
                }
            }
            // continue the rest of the normal stuff
            writer.WriteLine("          </KeyCodes>" +
                    "\r\n          <Context />" +
                    "\r\n          <X>0</X>" +
                    "\r\n          <Y>0</Y>" +
                    "\r\n          <Z>0</Z>" +
                    "\r\n          <InputMode>0</InputMode>" +
                    "\r\n          <ConditionPairing>0</ConditionPairing>" +
                    "\r\n          <ConditionGroup>0</ConditionGroup>" +
                    "\r\n          <ConditionStartOperator>0</ConditionStartOperator>" +
                    "\r\n          <ConditionStartValue>0</ConditionStartValue>" +
                    "\r\n          <ConditionStartValueType>0</ConditionStartValueType>" +
                    "\r\n          <ConditionStartType>0</ConditionStartType>" +
                    "\r\n          <DecimalContext1>0</DecimalContext1>" +
                    "\r\n          <DecimalContext2>0</DecimalContext2>" +
                    "\r\n          <DateContext1>0001-01-01T00:00:00</DateContext1>" +
                    "\r\n          <DateContext2>0001-01-01T00:00:00</DateContext2>" +
                    "\r\n          <Disabled>false</Disabled>" +
                    "\r\n          <RandomSounds />" +
                    "\r\n          <IntegerContext1>0</IntegerContext1>" +
                    "\r\n          <IntegerContext2>0</IntegerContext2>" +
                    "\r\n        </CommandAction>" +
                    "\r\n      </ActionSequence>");

        }

        // at this point the VA code is similar for commands that have keybinds and those that do not
        writer.WriteLine("      <Async>false</Async>" +
    "\r\n      <Enabled>true</Enabled>" +
    "\r\n      <UseShortcut>false</UseShortcut>" +
    "\r\n      <keyValue>0</keyValue>" +
    "\r\n      <keyShift>0</keyShift>" +
    "\r\n      <keyAlt>0</keyAlt>" +
    "\r\n      <keyCtrl>0</keyCtrl>" +
    "\r\n      <keyWin>0</keyWin>" +
    "\r\n      <keyPassthru>true</keyPassthru>" +
    "\r\n      <UseSpokenPhrase>true</UseSpokenPhrase>" +
    "\r\n      <onlyKeyUp>false</onlyKeyUp>" +
    "\r\n      <RepeatNumber>2</RepeatNumber>" +
    "\r\n      <RepeatType>0</RepeatType>" +
    "\r\n      <CommandType>0</CommandType>" +
    "\r\n      <SourceProfile>00000000-0000-0000-0000-000000000000</SourceProfile>" +
    "\r\n      <UseConfidence>false</UseConfidence>" +
    "\r\n      <minimumConfidenceLevel>0</minimumConfidenceLevel>" +
    "\r\n      <UseJoystick>false</UseJoystick>" +
    "\r\n      <joystickNumber>0</joystickNumber>" +
    "\r\n      <joystickButton>0</joystickButton>" +
    "\r\n      <joystickNumber2>0</joystickNumber2>" +
    "\r\n      <joystickButton2>0</joystickButton2>" +
    "\r\n      <joystickUp>false</joystickUp>" +
    "\r\n      <KeepRepeating>false</KeepRepeating>" +
    "\r\n      <UseProcessOverride>false</UseProcessOverride>" +
    "\r\n      <ProcessOverrideActiveWindow>true</ProcessOverrideActiveWindow>" +
    "\r\n      <LostFocusStop>false</LostFocusStop>" +
    "\r\n      <PauseLostFocus>false</PauseLostFocus>" +
    "\r\n      <LostFocusBackCompat>true</LostFocusBackCompat>" +
    "\r\n      <UseMouse>false</UseMouse>" +
    "\r\n      <Mouse1>false</Mouse1>" +
    "\r\n      <Mouse2>false</Mouse2>" +
    "\r\n      <Mouse3>false</Mouse3>" +
    "\r\n      <Mouse4>false</Mouse4>" +
    "\r\n      <Mouse5>false</Mouse5>" +
    "\r\n      <Mouse6>false</Mouse6>" +
    "\r\n      <Mouse7>false</Mouse7>" +
    "\r\n      <Mouse8>false</Mouse8>" +
    "\r\n      <Mouse9>false</Mouse9>" +
    "\r\n      <MouseUpOnly>false</MouseUpOnly>" +
    "\r\n      <MousePassThru>true</MousePassThru>" +
    "\r\n      <joystickExclusive>false</joystickExclusive>" +
    "\r\n      <lastEditedAction>135cd66d-2278-4090-b18c-d5d01c58f6f2</lastEditedAction>" + // does this have any effect?
    "\r\n      <UseProfileProcessOverride>false</UseProfileProcessOverride>" +
    "\r\n      <ProfileProcessOverrideActiveWindow>false</ProfileProcessOverrideActiveWindow>" +
    "\r\n      <RepeatIfKeysDown>false</RepeatIfKeysDown>" +
    "\r\n      <RepeatIfMouseDown>false</RepeatIfMouseDown>" +
    "\r\n      <RepeatIfJoystickDown>false</RepeatIfJoystickDown>" +
    "\r\n      <AH>0</AH>" +
    "\r\n      <CL>0</CL>" +
    "\r\n      <HasMB>false</HasMB>" +
    "\r\n      <UseVariableHotkey>false</UseVariableHotkey>" +
    "\r\n      <CLE>0</CLE>" +
    "\r\n      <EX1>false</EX1>" +
    "\r\n      <EX2>false</EX2>" +
    "\r\n      <InternalId xsi:nil=\"true\" />");

        // if there was a keybind, this is true, else, false
        if (currentKeycombo == "")// if there are no keybinds associated
        {
            writer.WriteLine("      <HasInput>false</HasInput>");
        }
        else
        {
            writer.WriteLine("      <HasInput>true</HasInput>");
        }

        writer.WriteLine("      <HotkeyDoubleTapLevel>0</HotkeyDoubleTapLevel>" +
        "\r\n      <MouseDoubleTapLevel>0</MouseDoubleTapLevel>" +
        "\r\n      <JoystickDoubleTapLevel>0</JoystickDoubleTapLevel>" +
        "\r\n      <HotkeyLongTapLevel>0</HotkeyLongTapLevel>" +
        "\r\n      <MouseLongTapLevel>0</MouseLongTapLevel>" +
        "\r\n      <JoystickLongTapLevel>0</JoystickLongTapLevel>" +
        "\r\n      <AlwaysExec>false</AlwaysExec>" +
        "\r\n      <ResourceBalance>0</ResourceBalance>" +
        "\r\n      <PreventExec>false</PreventExec>" +
        "\r\n      <ExternalEventsEnabled>false</ExternalEventsEnabled>" +
        "\r\n      <ExcludeExecOnRecognized>false</ExcludeExecOnRecognized>" +
        "\r\n      <UseVariableMouseShortcut>false</UseVariableMouseShortcut>" +
        "\r\n      <UseVariableJoystickShortcut>false</UseVariableJoystickShortcut>" +
        "\r\n    </Command>");
        //}
        commandCreationCounter++; // increase the command counter by one
    }

    // add the typical suffix to the end of the VA profile
    writer.WriteLine(vaSuffix);
}

// time to let the user know what happened
Console.WriteLine("Export complete.");
Console.WriteLine("Commands created: " + commandCreationCounter);
Console.WriteLine("Exported to: " + fullPath);
Console.WriteLine("Press Enter to quit.");
Console.ReadLine();

/* The GetKeycode method takes a string (or rather a char) snf
 * finds the resulting number (int in this case) to send back to
 * the section that tells VA what buttons are pressed in the profile
 * at the same time. All numbers/conversions were gotten manually
 * by making profiles and inspecting the XML file for which
 * buttons VA considered them to be. The name of the keyBind
 * comes from the DCS html.
 */
int GetKeycode(string keyBind)
{
    int keyCode;
    if (string.IsNullOrEmpty(keyBind))
    {
        keyCode = 0;
    }
    else if (keyBind.Equals("LCtrl"))
    {
        keyCode = 162;
    }
    else if (keyBind.Equals("RCtrl"))
    {
        keyCode = 163;
    }
    else if (keyBind.Equals("LShift"))
    {
        keyCode = 160;
    }
    else if (keyBind.Equals("RShift"))
    {
        keyCode = 161;
    }
    else if (keyBind.Equals("LAlt"))
    {
        keyCode = 164;
    }
    else if (keyBind.Equals("RAlt"))
    {
        keyCode = 165;
    }
    else if (keyBind.Equals("LWin"))
    {
        keyCode = 91;
    }
    else if (keyBind.Equals("RWin"))
    {
        keyCode = 92;
    }
    else if (keyBind.Equals("Up"))
    {
        keyCode = 38;
    }
    else if (keyBind.Equals("Down"))
    {
        keyCode = 40;
    }
    else if (keyBind.Equals("Left"))
    {
        keyCode = 37;
    }
    else if (keyBind.Equals("Right"))
    {
        keyCode = 39;
    }
    else if (keyBind.Equals("0"))
    {
        keyCode = 48;
    }
    else if (keyBind.Equals("1"))
    {
        keyCode = 49;
    }
    else if (keyBind.Equals("2"))
    {
        keyCode = 50;
    }
    else if (keyBind.Equals("3"))
    {
        keyCode = 51;
    }
    else if (keyBind.Equals("4"))
    {
        keyCode = 52;
    }
    else if (keyBind.Equals("5"))
    {
        keyCode = 53;
    }
    else if (keyBind.Equals("6"))
    {
        keyCode = 54;
    }
    else if (keyBind.Equals("7"))
    {
        keyCode = 55;
    }
    else if (keyBind.Equals("8"))
    {
        keyCode = 56;
    }
    else if (keyBind.Equals("9"))
    {
        keyCode = 57;
    }
    else if (keyBind.Equals("Enter"))
    {
        keyCode = 13;
    }
    else if (keyBind.Equals("NumEnter"))
    {
        keyCode = 156;
    }
    else if (keyBind.Equals("Esc")) // TODO Confirm 
    {
        keyCode = 27;
    }
    else if (keyBind.Equals("Space"))
    {
        keyCode = 32;
    }
    else if (keyBind.Equals("F1"))
    {
        keyCode = 112;
    }
    else if (keyBind.Equals("F2"))
    {
        keyCode = 113;
    }
    else if (keyBind.Equals("F3"))
    {
        keyCode = 114;
    }
    else if (keyBind.Equals("F4"))
    {
        keyCode = 115;
    }
    else if (keyBind.Equals("F5"))
    {
        keyCode = 116;
    }
    else if (keyBind.Equals("F6"))
    {
        keyCode = 117;
    }
    else if (keyBind.Equals("F7"))
    {
        keyCode = 118;
    }
    else if (keyBind.Equals("F8"))
    {
        keyCode = 119;
    }
    else if (keyBind.Equals("F9"))
    {
        keyCode = 120;
    }
    else if (keyBind.Equals("F10"))
    {
        keyCode = 121;
    }
    else if (keyBind.Equals("F11"))
    {
        keyCode = 122;
    }
    else if (keyBind.Equals("F12"))
    {
        keyCode = 123;
    }
    else if (keyBind.Equals("F13"))
    {
        keyCode = 124;
    }
    else if (keyBind.Equals("F14"))
    {
        keyCode = 125;
    }
    else if (keyBind.Equals("F15"))
    {
        keyCode = 126;
    }
    else if (keyBind.Equals("F16"))
    {
        keyCode = 127;
    }
    else if (keyBind.Equals("F17"))
    {
        keyCode = 128;
    }
    else if (keyBind.Equals("F18"))
    {
        keyCode = 129;
    }
    else if (keyBind.Equals("F19"))
    {
        keyCode = 130;
    }
    else if (keyBind.Equals("F20"))
    {
        keyCode = 131;
    }
    else if (keyBind.Equals("F21"))
    {
        keyCode = 132;
    }
    else if (keyBind.Equals("F22"))
    {
        keyCode = 133;
    }
    else if (keyBind.Equals("F23"))
    {
        keyCode = 134;
    }
    else if (keyBind.Equals("F24"))
    {
        keyCode = 135;
    }
    else if (keyBind.Equals("`"))
    {
        keyCode = 192;
    }
    else if (keyBind.Equals("Tab"))
    {
        keyCode = 9;
    }
    else if (keyBind.Equals("CapsLock"))
    {
        keyCode = 20;
    }
    else if (keyBind.Equals("Back"))
    {
        keyCode = 8;
    }
    else if (keyBind.Equals("="))
    {
        keyCode = 187;
    }
    else if (keyBind.Equals("-"))
    {
        keyCode = 189;
    }
    else if (keyBind.Equals("\\"))
    {
        keyCode = 220;
    }
    else if (keyBind.Equals("]"))
    {
        keyCode = 221;
    }
    else if (keyBind.Equals("["))
    {
        keyCode = 219;
    }
    else if (keyBind.Equals("'"))
    {
        keyCode = 222;
    }
    else if (keyBind.Equals(";"))
    {
        keyCode = 186;
    }
    else if (keyBind.Equals("/"))
    {
        keyCode = 191;
    }
    else if (keyBind.Equals("."))
    {
        keyCode = 190;
    }
    // this is like this to cover an edge case with csv import
    // TODO evaluate if required for html and tsv
    else if (keyBind.Equals("\",\"") || keyBind.Equals(",")) 
    {
        keyCode = 188;
    }
    else if (keyBind.Equals("NumLock"))
    {
        keyCode = 144;
    }
    else if (keyBind.Equals("Num."))
    {
        keyCode = 144; // TODO confirm
    }
    else if (keyBind.Equals("Num/"))
    {
        keyCode = 111;
    }
    else if (keyBind.Equals("Num*"))
    {
        keyCode = 106;
    }
    else if (keyBind.Equals("Num-"))
    {
        keyCode = 109;
    }
    else if (keyBind.Equals("Num+"))
    {
        keyCode = 107;
    }
    else if (keyBind.Equals("NumEnter"))
    {
        keyCode = 156;
    }
    else if (keyBind.Equals("Num."))
    {
        keyCode = 110;
    }
    else if (keyBind.Equals("Num0"))
    {
        keyCode = 96;
    }
    else if (keyBind.Equals("Num1"))
    {
        keyCode = 97;
    }
    else if (keyBind.Equals("Num2"))
    {
        keyCode = 98;
    }
    else if (keyBind.Equals("Num3"))
    {
        keyCode = 99;
    }
    else if (keyBind.Equals("Num4"))
    {
        keyCode = 100;
    }
    else if (keyBind.Equals("Num5"))
    {
        keyCode = 101;
    }
    else if (keyBind.Equals("Num6"))
    {
        keyCode = 102;
    }
    else if (keyBind.Equals("Num7"))
    {
        keyCode = 103;
    }
    else if (keyBind.Equals("Num8"))
    {
        keyCode = 104;
    }
    else if (keyBind.Equals("Num9"))
    {
        keyCode = 105;
    }
    else if (keyBind.Equals("PageDown"))
    {
        keyCode = 34;
    }
    else if (keyBind.Equals("End"))
    {
        keyCode = 35;
    }
    else if (keyBind.Equals("PageUp"))
    {
        keyCode = 33;
    }
    else if (keyBind.Equals("Home"))
    {
        keyCode = 36;
    }
    else if (keyBind.Equals("Delete"))
    {
        keyCode = 46;
    }
    else if (keyBind.Equals("Scroll"))
    {
        keyCode = 145;
    }
    else if (keyBind.Equals("Pause"))
    {
        keyCode = 19;
    }
    else if (keyBind.Equals("Insert"))
    {
        keyCode = 45;
    }
    else if (keyBind.Equals("Print")) // TODO exists in DCS?
    {
        keyCode = 44;
    }
    else if (keyBind.Equals("Esc")) // TODO exists in DCS?
    {
        keyCode = 27;
    }
    else if (keyBind.Equals("Insert"))
    {
        keyCode = 45;
    }
    else if (keyBind.Equals("Insert"))
    {
        keyCode = 45;
    }
    // alphabet
    else if (keyBind.Equals("A"))
    {
        keyCode = 65;
    }
    else if (keyBind.Equals("B"))
    {
        keyCode = 66;
    }
    else if (keyBind.Equals("C"))
    {
        keyCode = 67;
    }
    else if (keyBind.Equals("D"))
    {
        keyCode = 68;
    }
    else if (keyBind.Equals("E"))
    {
        keyCode = 69;
    }
    else if (keyBind.Equals("F"))
    {
        keyCode = 70;
    }
    else if (keyBind.Equals("G"))
    {
        keyCode = 71;
    }
    else if (keyBind.Equals("H"))
    {
        keyCode = 72;
    }
    else if (keyBind.Equals("I"))
    {
        keyCode = 73;
    }
    else if (keyBind.Equals("J"))
    {
        keyCode = 74;
    }
    else if (keyBind.Equals("K"))
    {
        keyCode = 75;
    }
    else if (keyBind.Equals("L"))
    {
        keyCode = 76;
    }
    else if (keyBind.Equals("M"))
    {
        keyCode = 77;
    }
    else if (keyBind.Equals("N"))
    {
        keyCode = 78;
    }
    else if (keyBind.Equals("O"))
    {
        keyCode = 79;
    }
    else if (keyBind.Equals("P"))
    {
        keyCode = 80;
    }
    else if (keyBind.Equals("Q"))
    {
        keyCode = 81;
    }
    else if (keyBind.Equals("R"))
    {
        keyCode = 82;
    }
    else if (keyBind.Equals("S"))
    {
        keyCode = 83;
    }
    else if (keyBind.Equals("T"))
    {
        keyCode = 84;
    }
    else if (keyBind.Equals("U"))
    {
        keyCode = 85;
    }
    else if (keyBind.Equals("V"))
    {
        keyCode = 86;
    }
    else if (keyBind.Equals("W"))
    {
        keyCode = 87;
    }
    else if (keyBind.Equals("X"))
    {
        keyCode = 88;
    }
    else if (keyBind.Equals("Y"))
    {
        keyCode = 89;
    }
    else if (keyBind.Equals("Z"))
    {
        keyCode = 90;
    }
    // end
    // this will typically result in "undefined" in the VA profile action description
    // which is good because that means the user can identify and debug
    else
    {
        keyCode = 0;
    }
    return keyCode;
}