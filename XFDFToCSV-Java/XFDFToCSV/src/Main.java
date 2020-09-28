import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.util.regex.Pattern;

public class Main {

	public static void main(String[] args) throws IOException {

		// Fields names and titles (csv) array
		String[] fieldsNames = { "02_PositionTitle", "03_Command", "04_Community", "05_Location", "06_PayPlanGrade",
				"07_MemberAcquisition", "08_JobSeries", "09_SupervisoryStatus", "10_FiscalYear", "11_YNCurrentLP",
				"12_CurrentLP", "13_OtherCurrentLP", "14_FurtherDevelopment", "15_YNInterestLP", "16_InterestLP",
				"17_OtherInterestLP", "01_ShortTermGoals", "02_LongTermGoals", "01_TechDevReq1", "02_TechDevReq2",
				"03_TechDevReq3", "04_TechDevReq4", "05_NonTechCompStrength1", "06_NonTechCompStrength2",
				"07_NonTechCompStrength3", "08_NonTechCompStrength4", "09_NonTechCompStrength5",
				"10_NonTechCompAreaGrow1", "11_NonTechCompAreaGrow2", "12_NonTechCompAreaGrow3",
				"13_NonTechCompAreaGrow4", "14_NonTechCompAreaGrow5", "01_Activity1Description", "02_Activity1DevType",
				"03_Activity1DevCategory", "04_Activity1PrimLearning", "05_Activity1SecLearning",
				"06_Activity1NonTechPrimComp", "07_Activity1NonTechSecComp", "08_Activity1EstTuitionCost",
				"09_Activity1EstTravelCost", "10_Activity1EstCompleteDate", "11_Activity1ActualCompleteDate",
				"12_Activity1OtherDevType", "13_Activity2Description", "14_Activity2DevType", "15_Activity2DevCategory",
				"16_Activity2PrimLearning", "17_Activity2SecLearning", "18_Activity2NonTechPrimComp",
				"19_Activity2NonTechSecComp", "20_Activity2EstTuitionCost", "21_Activity2EstTravelCost",
				"22_Activity2EstCompleteDate", "23_Activity2ActualCompleteDate", "24_Activity2OtherDevType",
				"25_Activity3Description", "26_Activity3DevType", "27_Activity3DevCategory", "28_Activity3PrimLearning",
				"29_Activity3SecLearning", "30_Activity3NonTechPrimComp", "31_Activity3NonTechSecComp",
				"32_Activity3EstTuitionCost", "33_Activity3EstTravelCost", "34_Activity3EstCompleteDate",
				"35_Activity3ActualCompleteDate", "36_Activity3OtherDevType", "37_Activity4Description",
				"38_Activity4DevType", "39_Activity4DevCategory", "40_Activity4PrimLearning", "41_Activity4SecLearning",
				"42_Activity4NonTechPrimComp", "43_Activity4NonTechSecComp", "44_Activity4EstTuitionCost",
				"45_Activity4EstTravelCost", "46_Activity4EstCompleteDate", "47_Activity4ActualCompleteDate",
				"48_Activity4OtherDevType", "01_IDPActivity1", "02_IDPActivity2", "03_IDPActivity3", "04_IDPActivity4",
				"05_IDPActivity5", "06_AccountabilityTactic1", "07_AccountabilityTactic2", "08_AccountabilityTactic3",
				"09_AccountabilityTactic4", "10_AccountabilityTactic5", "03_EmployeeDate", "04_IDPNotes",
				"06_SupervisorDate", "01_Activity5Description", "02_Activity5DevType", "03_Activity5DevCategory",
				"04_Activity5PrimLearning", "05_Activity5SecLearning", "06_Activity5NonTechPrimComp",
				"07_Activity5NonTechSecComp", "08_Activity5EstTuitionCost", "09_Activity5EstTravelCost",
				"10_Activity5EstCompleteDate", "11_Activity5ActualCompleteDate", "12_Activity5OtherDevType",
				"13_Activity6Description", "14_Activity6DevType", "15_Activity6DevCategory", "16_Activity6PrimLearning",
				"17_Activity6SecLearning", "18_Activity6NonTechPrimComp", "19_Activity6NonTechSecComp",
				"20_Activity6EstTuitionCost", "21_Activity6EstTravelCost", "22_Activity6EstCompleteDate",
				"23_Activity6ActualCompleteDate", "24_Activity6OtherDevType", "25_Activity7Description",
				"26_Activity7DevType", "27_Activity7DevCategory", "28_Activity7PrimLearning", "29_Activity7SecLearning",
				"30_Activity7NonTechPrimComp", "31_Activity7NonTechSecComp", "32_Activity7EstTuitionCost",
				"33_Activity7EstTravelCost", "34_Activity7EstCompleteDate", "35_Activity7ActualCompleteDate",
				"36_Activity7OtherDevType" };
		String[] fieldsTitles = { "1.02_PositionTitle", "1.03_Command", "1.04_Community", "1.05_Location",
				"1.06_PayPlanGrade", "1.07_MemberAcquisition", "1.08_JobSeries", "1.09_SupervisoryStatus",
				"1.10_FiscalYear", "1.11_YNCurrentLP", "1.12_CurrentLP", "1.13_OtherCurrentLP",
				"1.14_FurtherDevelopment", "1.15_YNInterestLP", "1.16_InterestLP", "1.17_OtherInterestLP",
				"2.01_ShortTermGoals", "2.02_LongTermGoals", "4.01_TechDevReq1", "4.02_TechDevReq2", "4.03_TechDevReq3",
				"4.04_TechDevReq4", "4.05_NonTechCompStrength1", "4.06_NonTechCompStrength2",
				"4.07_NonTechCompStrength3", "4.08_NonTechCompStrength4", "4.09_NonTechCompStrength5",
				"4.10_NonTechCompAreaGrow1", "4.11_NonTechCompAreaGrow2", "4.12_NonTechCompAreaGrow3",
				"4.13_NonTechCompAreaGrow4", "4.14_NonTechCompAreaGrow5", "5.01_Activity1Description",
				"5.02_Activity1DevType", "5.03_Activity1DevCategory", "5.04_Activity1PrimLearning",
				"5.05_Activity1SecLearning", "5.06_Activity1NonTechPrimComp", "5.07_Activity1NonTechSecComp",
				"5.08_Activity1EstTuitionCost", "5.09_Activity1EstTravelCost", "5.10_Activity1EstCompleteDate",
				"5.11_Activity1ActualCompleteDate", "5.12_Activity1OtherDevType", "5.13_Activity2Description",
				"5.14_Activity2DevType", "5.15_Activity2DevCategory", "5.16_Activity2PrimLearning",
				"5.17_Activity2SecLearning", "5.18_Activity2NonTechPrimComp", "5.19_Activity2NonTechSecComp",
				"5.20_Activity2EstTuitionCost", "5.21_Activity2EstTravelCost", "5.22_Activity2EstCompleteDate",
				"5.23_Activity2ActualCompleteDate", "5.24_Activity2OtherDevType", "5.25_Activity3Description",
				"5.26_Activity3DevType", "5.27_Activity3DevCategory", "5.28_Activity3PrimLearning",
				"5.29_Activity3SecLearning", "5.30_Activity3NonTechPrimComp", "5.31_Activity3NonTechSecComp",
				"5.32_Activity3EstTuitionCost", "5.33_Activity3EstTravelCost", "5.34_Activity3EstCompleteDate",
				"5.35_Activity3ActualCompleteDate", "5.36_Activity3OtherDevType", "5.37_Activity4Description",
				"5.38_Activity4DevType", "5.39_Activity4DevCategory", "5.40_Activity4PrimLearning",
				"5.41_Activity4SecLearning", "5.42_Activity4NonTechPrimComp", "5.43_Activity4NonTechSecComp",
				"5.44_Activity4EstTuitionCost", "5.45_Activity4EstTravelCost", "5.46_Activity4EstCompleteDate",
				"5.47_Activity4ActualCompleteDate", "5.48_Activity4OtherDevType", "6.01_IDPActivity1",
				"6.02_IDPActivity2", "6.03_IDPActivity3", "6.04_IDPActivity4", "6.05_IDPActivity5",
				"6.06_AccountabilityTactic1", "6.07_AccountabilityTactic2", "6.08_AccountabilityTactic3",
				"6.09_AccountabilityTactic4", "6.10_AccountabilityTactic5", "7.03_EmployeeDate", "7.04_IDPNotes",
				"7.06_SupervisorDate", "8.01_Activity5Description", "8.02_Activity5DevType",
				"8.03_Activity5DevCategory", "8.04_Activity5PrimLearning", "8.05_Activity5SecLearning",
				"8.06_Activity5NonTechPrimComp", "8.07_Activity5NonTechSecComp", "8.08_Activity5EstTuitionCost",
				"8.09_Activity5EstTravelCost", "8.10_Activity5EstCompleteDate", "8.11_Activity5ActualCompleteDate",
				"8.12_Activity5OtherDevType", "8.13_Activity6Description", "8.14_Activity6DevType",
				"8.15_Activity6DevCategory", "8.16_Activity6PrimLearning", "8.17_Activity6SecLearning",
				"8.18_Activity6NonTechPrimComp", "8.19_Activity6NonTechSecComp", "8.20_Activity6EstTuitionCost",
				"8.21_Activity6EstTravelCost", "8.22_Activity6EstCompleteDate", "8.23_Activity6ActualCompleteDate",
				"8.24_Activity6OtherDevType", "8.25_Activity7Description", "8.26_Activity7DevType",
				"8.27_Activity7DevCategory", "8.28_Activity7PrimLearning", "8.29_Activity7SecLearning",
				"8.30_Activity7NonTechPrimComp", "8.31_Activity7NonTechSecComp", "8.32_Activity7EstTuitionCost",
				"8.33_Activity7EstTravelCost", "8.34_Activity7EstCompleteDate", "8.35_Activity7ActualCompleteDate",
				"8.36_Activity7OtherDevType" };

		// Get txt files folder path
		// Console.WriteLine("IDP Data Extraction Script - Last version 6/2/2020\n\nTo
		// start the process, please make sure all txt files are in one folder. When the
		// process is finished, this window will close and a CSV file called
		// 'IDP_final_results.csv' will be created in the same folder of the txt
		// files.\n\nPlease provide the txt folder path(eg.: E:/User/myfolder):");
		// String folderPath = Console.ReadLine();
		System.out.println(
				"IDP Data Extraction Script - Last version 6/2/2020\n\nTo start the process, please make sure all txt files are in one folder. When the process is finished, this window will close and a CSV file called 'IDP_final_results.csv' will be created in the same folder of the txt files.\n\nThe txt folder path is 'C:\\Users\\aaron.zeman1\\Desktop\\IDP'");
		// String folderPath = "C:\\Users\\aaron.zeman1\\Desktop\\IDP";
		String folderPath = "C:\\Users\\Hp\\Documents\\Guyot\\IDP PDF\\FilesTxt";

		// Set the csv delimiters
		String fieldDelimiter = ",";
		String rowDelimiter = "\n";

		// Create output CSV file
		String outputFileName = "IDP_final_results.csv";
		String outputFilePath = folderPath + "/" + outputFileName;
		FileWriter outputFile = new FileWriter(outputFilePath);
		String titles = String.join(fieldDelimiter, fieldsTitles) + rowDelimiter;
		outputFile.write(titles);
		System.out.println(titles);
		outputFile.close();

		// Iterate in directory
		String[] files = new File(folderPath).list();
		for (String fileName : files) {
			System.out.println(fileName);
			// all files except the output csv
			if (!fileName.equals(outputFileName)) {
				String filePath = folderPath + "/" + fileName;
				// Delete field delimiter (,) and replace (") with (')
				FileReader fileReader = new FileReader(filePath);
				BufferedReader bufferedReader = new BufferedReader(fileReader);
				String text = "";
				String aux;
				while ((aux = bufferedReader.readLine()) != null) {
					text += aux;
				}
				bufferedReader.close();
				fileReader.close();
//				System.out.println(text);
                text = text.replace('"', '\'').replace(fieldDelimiter,"");

                String res = "";

                //extract value from each field
                for (String label : fieldsNames)
                {
                    
//                    Pattern pattern = Pattern.compile("<field name='" + label + "'\n><value\n>(.*)</value\n></field\n>");
//                    if(pattern.matcher(text).matches()) {
//                    	System.out.println(label + " - YES");
//                    }
//                    
//                    //the pattern is different if the field is multiline
//                    boolean isMultilineField = label.equals("01_ShortTermGoals") || label.equals("02_LongTermGoals");
//                    if (isMultilineField)
//                    {
//                        pattern = @"<field name='" + label + @"'\n><value-richtext\n><body .*?\n>((.|\n)*?)<\/body\n><\/value-richtext\n><\/field\n>";
//                    }
//
//                    //get value between tags
//                    String val = Regex.Match(text, pattern).Groups[1].Value;
//
//
//                    //delete all "select option" label
//                    //TODO_ 2 options added (OBSERVED)
//                    String[] selectOptionLabels = { "-- Select Command --", "-- Select Community --", "-- Select your Location --", "-- Select Pay Plan/Grade --", "-- Select Option --", "-- Select Status --", "-- Select an option --", "-- Select Program --","--Select an option --" };
//                    for (String option : selectOptionLabels)
//                    {
//                        val = val.Equals(option) ? "" : val;
//                    }
//
//
//                    //get each paragraph value from multiline. Delete line breaks, so each paragraph start with * (csv limitations)
//                    String strToReplace1 = "<p dir='ltr' style='margin-top:0pt;margin-bottom:0pt;font-family:Helvetica;font-size:12pt'\n>";
//                    String strToReplace2 = "</p\n>";
//                    val = val.Replace(strToReplace1," *").Replace(strToReplace2," ").Replace("-"," -");
//
//            
//                    //append value to result line
//                    res += val; 
//                    if (!label.Equals(fieldsNames.GetValue(fieldsNames.Length - 1)))
//                    {
//                        res += fieldDelimiter;
//                    }
                }
//
//                //append row delimiter to result line
//                res += rowDelimiter;
//
//                //append result line to CSV file
//                File.AppendAllText(outputFilePath, res);
			}
		}
	}

}
