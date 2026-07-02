using System;
using System.IO;
using HXMSExcelPing;

class ExcelEvaluator
{
    static void Main(string[] args)
    {
        if (args.Length == 0)
        {
            Console.WriteLine("Usage: ExcelEvaluator.exe <question_id>");
            Console.WriteLine("Example: ExcelEvaluator.exe 68617");
            return;
        }

        string questionNo = args[0];
        string questionDir = Path.Combine("excel_questions", questionNo);

        if (!Directory.Exists(questionDir))
        {
            Console.WriteLine("Error: question directory does not exist - " + questionDir);
            return;
        }

        string configFile = Path.Combine(questionDir, "config.xml");
        if (!File.Exists(configFile))
        {
            Console.WriteLine("Error: scoring config file does not exist - " + configFile);
            return;
        }

        string inputDir = Path.Combine(questionDir, "input");
        if (!Directory.Exists(inputDir))
        {
            Console.WriteLine("Error: input directory does not exist - " + inputDir);
            return;
        }

        Console.WriteLine("======================================================================");
        Console.WriteLine("Excel task evaluation");
        Console.WriteLine("======================================================================");
        Console.WriteLine();
        Console.WriteLine("Question ID: " + questionNo);
        Console.WriteLine("Question directory: " + Path.GetFullPath(questionDir));
        Console.WriteLine("Input directory: " + Path.GetFullPath(inputDir));
        Console.WriteLine();
        Console.WriteLine("Starting evaluation...");
        Console.WriteLine();

        try
        {
            string ansCont = File.ReadAllText(configFile);

            MSExcelPing ping = new MSExcelPing();
            ping.AnsCont = ansCont;
            ping.TestPath = Path.GetFullPath(inputDir);
            ping.GetTotalScore = 20f;
            ping.AllGM = true;

            float score = ping.Ping();

            Console.WriteLine("======================================================================");
            Console.WriteLine("Evaluation result");
            Console.WriteLine("======================================================================");
            Console.WriteLine();

            if (score < 0)
            {
                Console.WriteLine("Score: " + score + " (evaluation error)");
            }
            else
            {
                Console.WriteLine("Score: " + score + " / 20");
                Console.WriteLine("Score rate: " + (score / 20 * 100).ToString("F1") + "%");
            }
            Console.WriteLine();

            if (!string.IsNullOrEmpty(ping.ShowErr))
            {
                Console.WriteLine("Error messages:");
                Console.WriteLine(ping.ShowErr);
                Console.WriteLine();
            }

            if (!string.IsNullOrEmpty(ping.GradingMessage))
            {
                Console.WriteLine("Scoring details:");
                Console.WriteLine(ping.GradingMessage);
                Console.WriteLine();
            }

            Console.WriteLine("======================================================================");

            // Save evaluation output.
            string resultFile = Path.Combine(questionDir, "evaluation_result.txt");
            using (StreamWriter sw = new StreamWriter(resultFile, false, System.Text.Encoding.UTF8))
            {
                sw.WriteLine("Question ID: " + questionNo);
                sw.WriteLine("Score: " + score + " / 20");
                sw.WriteLine("Score rate: " + (score / 20 * 100).ToString("F1") + "%");
                sw.WriteLine();
                if (!string.IsNullOrEmpty(ping.ShowErr))
                {
                    sw.WriteLine("Error messages:");
                    sw.WriteLine(ping.ShowErr);
                    sw.WriteLine();
                }
                if (!string.IsNullOrEmpty(ping.GradingMessage))
                {
                    sw.WriteLine("Scoring details:");
                    sw.WriteLine(ping.GradingMessage);
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Evaluation failed: " + ex.Message);
            Console.WriteLine(ex.StackTrace);
        }
    }
}
