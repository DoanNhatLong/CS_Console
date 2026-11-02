using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Data;
using ExcelDataReader;

class Program
{
    static void Main(string[] args)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // --- CẤU HÌNH THAM SỐ CHO FILE SK.xlsx ---
        string filePath = @"D:\0 WordPress\SK.xlsx"; // Đường dẫn tới file SK.xlsx của bạn
        string jsonOptionsPath = "quiz_options.json";

        if (!File.Exists(filePath))
        {
            Console.WriteLine($"[Lỗi] Không tìm thấy file Excel tại: {filePath}");
            return;
        }

        if (!File.Exists(jsonOptionsPath))
        {
            Console.WriteLine($"[Lỗi] Không tìm thấy file cấu hình '{jsonOptionsPath}' trong thư mục project!");
            return;
        }

        string quizOptionsJson = File.ReadAllText(jsonOptionsPath, Encoding.UTF8);

        string prefixTitle = "SKK5"; 
        int startingQuestionId = 18803; // ID câu hỏi bắt đầu theo yêu cầu
        int currentQuestionId = startingQuestionId;

        int authorId = 5;
        int quizCategoryId = 36; // Mã danh mục mới
        
        // Tên pack theo yêu cầu mới
        string quizTitlePrefix = "Saskatchewan Knowledge - Pack";

        List<string> questionSqls = new List<string>();
        List<string> answerSqls = new List<string>();
        List<int> allQuestionIds = new List<int>();

        Console.WriteLine($"Đang đọc file '{filePath}' (ID bắt đầu từ {startingQuestionId}, gom nhóm 20 câu/pack)...");

        using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
        {
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = false }
                });

                var table = result.Tables[0];
                int rowCount = table.Rows.Count;

                for (int i = 0; i < rowCount; i++)
                {
                    var row = table.Rows[i];

                    if (row[0] == DBNull.Value || row[1] == DBNull.Value) continue;

                    string sttStr = row[0].ToString().Trim();
                    if (!int.TryParse(sttStr, out int stt)) continue;

                    string questionText = row[1].ToString().Trim();

                    string ans1 = row[2] != DBNull.Value ? row[2].ToString().Trim() : "";
                    string ans2 = row[3] != DBNull.Value ? row[3].ToString().Trim() : "";
                    string ans3 = row[4] != DBNull.Value ? row[4].ToString().Trim() : "";
                    string ans4 = row[5] != DBNull.Value ? row[5].ToString().Trim() : "";
                    List<string> answers = new List<string> { ans1, ans2, ans3, ans4 };

                    string explanation = row[6] != DBNull.Value ? row[6].ToString().Trim() : "";
                    string correctAnswerLetter = row[7] != DBNull.Value ? row[7].ToString().Trim().ToUpper() : "A";
                    string categoryIdStr = row[8] != DBNull.Value ? row[8].ToString().Trim() : "0";

                    string questionTitle = $"{prefixTitle}{stt.ToString("D4")}";

                    // 1. Sinh câu lệnh SQL cho bảng questions
                    string qSql = $"INSERT INTO `snnj_aysquiz_questions` (`id`, `author_id`, `category_id`, `question`, `question_title`, `question_image`, `wrong_answer_text`, `right_answer_text`, `question_hint`, `explanation`, `type`, `published`, `create_date`, `not_influence_to_score`, `weight`, `options`, `user_explanation`) VALUES " +
                                  $"({currentQuestionId}, {authorId}, {categoryIdStr}, '{EscapeSql(questionText)}', '{questionTitle}', NULL, NULL, NULL, NULL, '{EscapeSql(explanation)}', 'radio', 1, NULL, NULL, 1, NULL, NULL);";
                    questionSqls.Add(qSql);

                    // 2. Sinh câu lệnh SQL cho bảng answers
                    for (int j = 0; j < answers.Count; j++)
                    {
                        char currentLetter = (char)('A' + j);
                        int isCorrect = (currentLetter.ToString() == correctAnswerLetter) ? 1 : 0;
                        int ordering = j + 1;

                        string aSql = $"INSERT INTO `snnj_aysquiz_answers` (`question_id`, `answer`, `image`, `correct`, `ordering`, `weight`, `placeholder`) VALUES " +
                                      $"({currentQuestionId}, '{EscapeSql(answers[j])}', NULL, {isCorrect}, {ordering}, 0, NULL);";
                        answerSqls.Add(aSql);
                    }

                    allQuestionIds.Add(currentQuestionId);
                    currentQuestionId++;
                }
            }
        }

        // 3. Tự động gom nhóm thành các Pack (mỗi pack tối đa 20 câu)
        List<string> quizSqls = new List<string>();
        int packSize = 20; 
        int totalPacks = (int)Math.Ceiling((double)allQuestionIds.Count / packSize);

        for (int p = 0; p < totalPacks; p++)
        {
            var chunkIds = allQuestionIds.GetRange(p * packSize, Math.Min(packSize, allQuestionIds.Count - (p * packSize)));
            string idsString = string.Join(",", chunkIds);
            
            string packName = $"{quizTitlePrefix} {(p + 1).ToString("D2")}";

            string quizSql = $"INSERT INTO `snnj_aysquiz_quizes` (`author_id`, `title`, `description`, `quiz_category_id`, `question_ids`, `ordering`, `published`, `options`) VALUES " +
                           $"({authorId}, '{EscapeSql(packName)}', '', {quizCategoryId}, '{idsString}', 0, 1, '{EscapeSql(quizOptionsJson)}');";
            quizSqls.Add(quizSql);
        }

        // Ghi ra 3 file SQL mới
        File.WriteAllLines("questions.sql", questionSqls, Encoding.UTF8);
        File.WriteAllLines("answers.sql", answerSqls, Encoding.UTF8);
        File.WriteAllLines("quizes.sql", quizSqls, Encoding.UTF8);

        Console.WriteLine($"[Thành công!] Đã xuất ra 3 file SQL từ file SK.xlsx:");
        Console.WriteLine($" - questions.sql ({questionSqls.Count} câu hỏi, ID từ {startingQuestionId})");
        Console.WriteLine($" - answers.sql ({answerSqls.Count} đáp án)");
        Console.WriteLine($" - quizes.sql ({quizSqls.Count} quiz pack, mỗi pack tối đa 20 câu)");
    }

    static string EscapeSql(string input)
    {
        if (string.IsNullOrEmpty(input)) return "";
        return input.Replace("'", "''");
    }
}