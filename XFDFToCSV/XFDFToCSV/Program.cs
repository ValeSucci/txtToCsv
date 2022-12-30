using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Web;
using System.Linq;

namespace XFDFToCSV
{
    class Program
    {
        static void Main(string[] args)
        {
            //Fields names and titles (csv) array
            /*v3 - 12-15-21 - new fields*/
            string[] fieldsNames = { "01_Name","00_Email","02_PositionTitle", "03_Command", "04_Community", "05_Location", "06_CareerSegment", "07_MemberAcquisition", "08_JobSeries", "09_SupervisoryStatus", "10_FiscalYear", "11_YNCurrentLP", "12_CurrentLP", "13_OtherCurrentLP", "14_FurtherDevelopment", "15_YNInterestLP", "16_InterestLP", "17_OtherInterestLP", "18_YNCurrentMentored", "19_MentoredDevPeriod", "20_YNCurrentlyMentoring", "21_YNMentoringDevPeriod", "22_YNRotational", "23_CommandRotate", "24_CommunityRotate", "25_OtherRotationFuture", "26_YNCoaching", "27_YNCurrentlyCoaching", "28_YNInterestedCoaching", "29_YNCurrentIDPUploaded", "01_ShortTermGoals", "02_LongTermGoals", "01_TechDevReq1", "02_TechDevReq2", "03_TechDevReq3", "04_TechDevReq4", "05_NonTechCompStrength1", "06_NonTechCompStrength2", "07_NonTechCompStrength3", "08_NonTechCompStrength4", "09_NonTechCompStrength5", "10_NonTechCompAreaGrow1", "11_NonTechCompAreaGrow2", "12_NonTechCompAreaGrow3", "13_NonTechCompAreaGrow4", "14_NonTechCompAreaGrow5", "01_Activity1Description", "02_Activity1DevType", "03_Activity1DevCategory", "04_Activity1PrimLearning", "05_Activity1SecLearning", "06_Activity1NonTechPrimComp", "07_Activity1NonTechSecComp", "08_Activity1EstTuitionCost", "09_Activity1EstTravelCost", "10_Activity1UnknownCostInfo", "11_Activity1EstCompleteDate", "12_Activity1ActualCompleteDate", "13_Activity1OtherDevType", "14_Activity2Description", "15_Activity2DevType", "16_Activity2DevCategory", "17_Activity2PrimLearning", "18_Activity2SecLearning", "19_Activity2NonTechPrimComp", "20_Activity2NonTechSecComp", "21_Activity2EstTuitionCost", "22_Activity2EstTravelCost", "23_Activity2UnknownCostInfo", "24_Activity2EstCompleteDate", "25_Activity2ActualCompleteDate", "26_Activity2OtherDevType", "27_Activity3Description", "28_Activity3DevType", "29_Activity3DevCategory", "30_Activity3PrimLearning", "31_Activity3SecLearning", "32_Activity3NonTechPrimComp", "33_Activity3NonTechSecComp", "34_Activity3EstTuitionCost", "35_Activity3EstTravelCost", "36_Activity3UnknownCostInfo", "37_Activity3EstCompleteDate", "38_Activity3ActualCompleteDate", "39_Activity3OtherDevType", "40_Activity4Description", "41_Activity4DevType", "42_Activity4DevCategory", "43_Activity4PrimLearning", "44_Activity4SecLearning", "45_Activity4NonTechPrimComp", "46_Activity4NonTechSecComp", "47_Activity4EstTuitionCost", "48_Activity4EstTravelCost", "49_Activity4UnknownCostInfo", "50_Activity4EstCompleteDate", "51_Activity4ActualCompleteDate", "52_Activity4OtherDevType", "01_IDPActivity1", "02_IDPActivity2", "03_IDPActivity3", "04_IDPActivity4", "05_IDPActivity5", "06_AccountabilityTactic1", "07_AccountabilityTactic2", "08_AccountabilityTactic3", "09_AccountabilityTactic4", "10_AccountabilityTactic5", "01_IDPNotes", "03_EmployeeDate", "06_SupervisorDate", "01_Activity5Description", "02_Activity5DevType", "03_Activity5DevCategory", "04_Activity5PrimLearning", "05_Activity5SecLearning", "06_Activity5NonTechPrimComp", "07_Activity5NonTechSecComp", "08_Activity5EstTuitionCost", "09_Activity5EstTravelCost", "10_Activity5EstCompleteDate", "11_Activity5ActualCompleteDate", "12_Activity5OtherDevType", "13_Activity6Description", "14_Activity6DevType", "15_Activity6DevCategory", "16_Activity6PrimLearning", "17_Activity6SecLearning", "18_Activity6NonTechPrimComp", "19_Activity6NonTechSecComp", "20_Activity6EstTuitionCost", "21_Activity6EstTravelCost", "22_Activity6EstCompleteDate", "23_Activity6ActualCompleteDate", "24_Activity6OtherDevType", "25_Activity7Description", "26_Activity7DevType", "27_Activity7DevCategory", "28_Activity7PrimLearning", "29_Activity7SecLearning", "30_Activity7NonTechPrimComp", "31_Activity7NonTechSecComp", "32_Activity7EstTuitionCost", "33_Activity7EstTravelCost", "34_Activity7EstCompleteDate", "35_Activity7ActualCompleteDate", "36_Activity7OtherDevType" };
            string[] fieldsTitles = { "1.01_Name","1.00_Email","1.02_PositionTitle", "1.03_Command", "1.04_Community", "1.05_Location", "1.06_CareerSegment", "1.07_MemberAcquisition", "1.08_JobSeries", "1.09_SupervisoryStatus", "1.10_FiscalYear", "1.11_YNCurrentLP", "1.12_CurrentLP", "1.13_OtherCurrentLP", "1.14_FurtherDevelopment", "1.15_YNInterestLP", "1.16_InterestLP", "1.17_OtherInterestLP", "1.18_YNCurrentMentored", "1.19_MentoredDevPeriod", "1.20_YNCurrentlyMentoring", "1.21_YNMentoringDevPeriod", "1.22_YNRotational", "1.23_CommandRotate", "1.24_CommunityRotate", "1.25_OtherRotationFuture", "1.26_YNCoaching", "1.27_YNCurrentlyCoaching", "1.28_YNInterestedCoaching", "1.29_YNCurrentIDPUploaded", "2.01_ShortTermGoals", "2.02_LongTermGoals", "4.01_TechDevReq1", "4.02_TechDevReq2", "4.03_TechDevReq3", "4.04_TechDevReq4", "4.05_NonTechCompStrength1", "4.06_NonTechCompStrength2", "4.07_NonTechCompStrength3", "4.08_NonTechCompStrength4", "4.09_NonTechCompStrength5", "4.10_NonTechCompAreaGrow1", "4.11_NonTechCompAreaGrow2", "4.12_NonTechCompAreaGrow3", "4.13_NonTechCompAreaGrow4", "4.14_NonTechCompAreaGrow5", "5.01_Activity1Description", "5.02_Activity1DevType", "5.03_Activity1DevCategory", "5.04_Activity1PrimLearning", "5.05_Activity1SecLearning", "5.06_Activity1NonTechPrimComp", "5.07_Activity1NonTechSecComp", "5.08_Activity1EstTuitionCost", "5.09_Activity1EstTravelCost", "5.10_Activity1UnknownCostInfo", "5.11_Activity1EstCompleteDate", "5.12_Activity1ActualCompleteDate", "5.13_Activity1OtherDevType", "5.14_Activity2Description", "5.15_Activity2DevType", "5.16_Activity2DevCategory", "5.17_Activity2PrimLearning", "5.18_Activity2SecLearning", "5.19_Activity2NonTechPrimComp", "5.20_Activity2NonTechSecComp", "5.21_Activity2EstTuitionCost", "5.22_Activity2EstTravelCost", "5.23_Activity2UnknownCostInfo", "5.24_Activity2EstCompleteDate", "5.25_Activity2ActualCompleteDate", "5.26_Activity2OtherDevType", "5.27_Activity3Description", "5.28_Activity3DevType", "5.29_Activity3DevCategory", "5.30_Activity3PrimLearning", "5.31_Activity3SecLearning", "5.32_Activity3NonTechPrimComp", "5.33_Activity3NonTechSecComp", "5.34_Activity3EstTuitionCost", "5.35_Activity3EstTravelCost", "5.36_Activity3UnknownCostInfo", "5.37_Activity3EstCompleteDate", "5.38_Activity3ActualCompleteDate", "5.39_Activity3OtherDevType", "5.40_Activity4Description", "5.41_Activity4DevType", "5.42_Activity4DevCategory", "5.43_Activity4PrimLearning", "5.44_Activity4SecLearning", "5.45_Activity4NonTechPrimComp", "5.46_Activity4NonTechSecComp", "5.47_Activity4EstTuitionCost", "5.48_Activity4EstTravelCost", "5.49_Activity4UnknownCostInfo", "5.50_Activity4EstCompleteDate", "5.51_Activity4ActualCompleteDate", "5.52_Activity4OtherDevType", "6.01_IDPActivity1", "6.02_IDPActivity2", "6.03_IDPActivity3", "6.04_IDPActivity4", "6.05_IDPActivity5", "6.06_AccountabilityTactic1", "6.07_AccountabilityTactic2", "6.08_AccountabilityTactic3", "6.09_AccountabilityTactic4", "6.10_AccountabilityTactic5", "7.01_IDPNotes", "7.03_EmployeeDate", "7.06_SupervisorDate", "8.01_Activity5Description", "8.02_Activity5DevType", "8.03_Activity5DevCategory", "8.04_Activity5PrimLearning", "8.05_Activity5SecLearning", "8.06_Activity5NonTechPrimComp", "8.07_Activity5NonTechSecComp", "8.08_Activity5EstTuitionCost", "8.09_Activity5EstTravelCost", "8.10_Activity5EstCompleteDate", "8.11_Activity5ActualCompleteDate", "8.12_Activity5OtherDevType", "8.13_Activity6Description", "8.14_Activity6DevType", "8.15_Activity6DevCategory", "8.16_Activity6PrimLearning", "8.17_Activity6SecLearning", "8.18_Activity6NonTechPrimComp", "8.19_Activity6NonTechSecComp", "8.20_Activity6EstTuitionCost", "8.21_Activity6EstTravelCost", "8.22_Activity6EstCompleteDate", "8.23_Activity6ActualCompleteDate", "8.24_Activity6OtherDevType", "8.25_Activity7Description", "8.26_Activity7DevType", "8.27_Activity7DevCategory", "8.28_Activity7PrimLearning", "8.29_Activity7SecLearning", "8.30_Activity7NonTechPrimComp", "8.31_Activity7NonTechSecComp", "8.32_Activity7EstTuitionCost", "8.33_Activity7EstTravelCost", "8.34_Activity7EstCompleteDate", "8.35_Activity7ActualCompleteDate", "8.36_Activity7OtherDevType" };

            //Get txt files folder path
            //Console.WriteLine("IDP Data Extraction Script - Last version 6/2/2020\n\nTo start the process, please make sure all txt files are in one folder. When the process is finished, this window will close and a CSV file called 'IDP_final_results.csv' will be created in the same folder of the txt files.\n\nPlease provide the txt folder path(eg.: E:/User/myfolder):");
            //string folderPath = Console.ReadLine();

            //string folderPath = "C:\\Users\\aakzeman\\Desktop\\IDP"; //Aaron
            string folderPath = "C:\\Users\\valiaga\\Desktop\\TestingIDP"; //Vale A
            //string folderPath = "D:\\ETTE\\IDP PDF\\FilesTxt3"; //Vale S
            //Console.WriteLine("IDP Data Extraction Script - Last version 17/2/2020\n\nTo start the process, please make sure all txt files are in one folder. When the process is finished, this window will close and a CSV file called 'IDP_final_results.csv' will be created in the same folder of the txt files.\n\nThe txt folder path is '"+folderPath+"'");
            //v2 - 04-27-21 - date
            //Console.WriteLine("IDP Data Extraction Script - Last version 04/27/2021\n\nTo start the process, please make sure all txt files are in one folder. When the process is finished, this window will close and a CSV file called 'IDP_final_results.csv' will be created in the same folder of the txt files.\n\nThe txt folder path is '" + folderPath + "'");
            /*v3.1 - 12-30-22 - aaron path changes*/
            Console.WriteLine("IDP Data Extraction Script - Last version 12/30/2022\n\nTo start the process, please make sure all txt files are in one folder. When the process is finished, this window will close and a CSV file called 'IDP_final_results.csv' will be created in the same folder of the txt files.\n\nThe txt folder path is '" + folderPath + "'");

            //Set the csv delimiters
            string fieldDelimiter = ",";
            string rowDelimiter = "\n";

            //Create output CSV file
            string outputFileName = "IDP_final_results.csv";
            string outputFilePath = folderPath + "/"+outputFileName;
            File.WriteAllText(outputFilePath, string.Join(fieldDelimiter,fieldsTitles)+rowDelimiter);

            //Iterate in directory
            string[] files = Directory.GetFiles(folderPath);
            foreach (string filePath in files) {
                //all files except the output csv
                if (!filePath.Equals(folderPath+"\\"+outputFileName)) {
                    //Delete field delimiter (,) and replace (") with (')
                    string text = File.ReadAllText(filePath);
                    text = text.Replace('"', '\'').Replace(fieldDelimiter,"");

                    string res = "";

                    //extract value from each field
                    foreach (string label in fieldsNames)
                    {
            
                        string pattern = @"<field name='" + label + @"'\r?\n><value\r?\n>(.*)<\/value\r?\n><\/field\r?\n>";

                        //the pattern is different if the field is multiline
                        bool isMultilineField = label.Equals("01_ShortTermGoals") || label.Equals("02_LongTermGoals");
                        if (isMultilineField)
                        {
                            pattern = @"<field name='" + label + @"'\r?\n><value-richtext\r?\n><body .*?\r?\n>((.|\r?\n)*?)<\/body\r?\n><\/value-richtext\r?\n><\/field\r?\n>";
                        } else if (label.Equals("01_IDPNotes"))
                        {
                            //different multiline
                            pattern = @"<field name='" + label + @"'\r?\n><value\r?\n>((.|\r?\n)*?)<\/value\r?\n><\/field\r?\n>";
                        }

                        //get value between tags (each field)
                        string val = Regex.Match(text, pattern).Groups[1].Value;


                        //delete all "select option" label
                        //TODO_ 2 options added (OBSERVED)
                        string[] selectOptionLabels = { "-- Select Command --", "-- Select Community --", "-- Select your Location --", "-- Select Pay Plan/Grade --", "-- Select Option --", "-- Select Status --", "-- Select an option --", "-- Select Program --","-- Select Career Segment --","--Select an option --" };
                        foreach (string option in selectOptionLabels)
                        {
                            val = val.Equals(option) ? "" : val;
                        }
                        //validate numeric fields (fieldNames)
                        string[] numericFields = { "08_Activity1EstTuitionCost", "09_Activity1EstTravelCost", "21_Activity2EstTuitionCost", "22_Activity2EstTravelCost", "34_Activity3EstTuitionCost", "35_Activity3EstTravelCost", "47_Activity4EstTuitionCost", "48_Activity4EstTravelCost" };
                        if (numericFields.Contains(label)) 
                        {
                            int auxInt;
                            if (int.TryParse(val, out auxInt))
                            {
                                val = (auxInt > 50000) ? "50000" : val;
                            } else
                            {
                                val = "";
                            }                        
                        }
                        

                        //get each paragraph value from multiline. Delete line breaks, so each paragraph start with * (csv limitations)
                        if (isMultilineField || label.Equals("01_IDPNotes"))
                        {
                            string aux = val;
                            val = "";
                            string pattern2 = isMultilineField ? @"<p .*\r?\n>(.*)</p\r?\n>" : @"([^\r\n]+)";
                            MatchCollection paragraphs = Regex.Matches(aux, pattern2);
                            foreach (Match p in paragraphs)
                            {
                                val += " *" + p.Groups[1].Value;
                            }

                        }

               
                        //adding a blank space before "-" if it is at the beginning of a word (csv limitations) 
                        val = Regex.Replace(val, @"\B-", " -");

                        //decoding special characters like "&"
                        val = HttpUtility.HtmlDecode(val);


                        //append value to result line
                        res += val; 
                        if (!label.Equals(fieldsNames.GetValue(fieldsNames.Length - 1)))
                        {
                            res += fieldDelimiter;
                        }
                    }

                    //append row delimiter to result line
                    res += rowDelimiter;

                    //append result line to CSV file
                    File.AppendAllText(outputFilePath, res);
                }
            }                      
        }
    }

    /*
     * Eliminar las , o ; del texto (pdf) para mantener bien el csv -> matar las comas
     * Saltos de línea en un mismo campo pueden ser eliminados o suplantados por algo (eg. |) -> * inicio
     * Delimitador de la fila actual es salto de línea
     * 
     * El csv sale al mismo directorio (cambiar nombre o file?) -> IDP_final_results.csv
     * Editar texto inicial? para q ingresen el directorio -> Vale lo redacta
     * 
     */
}
