using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ClosedXML.Excel; // Thêm namespace ClosedXML
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

class Program
{
    static async Task Main(string[] args)
    {
        // Đặt mã hóa console thành UTF-8 để hiển thị tiếng Việt đúng
        Console.OutputEncoding = System.Text.Encoding.UTF8;

        // Nhận đầu vào đường dẫn thư mục từ người dùng
        Console.WriteLine("Nhập đường dẫn thư mục cần quét:");
        string rootDirectory = Console.ReadLine(); // Đọc đường dẫn từ console

        // Kiểm tra nếu đường dẫn không hợp lệ
        if (!Directory.Exists(rootDirectory))
        {
            Console.WriteLine("Đường dẫn không hợp lệ! Vui lòng kiểm tra lại.");
            return;
        }

        List<FileInfo> pdfFiles = new List<FileInfo>();

        // Quét tất cả các file PDF trong thư mục và các thư mục con
        GetPdfFiles(rootDirectory, pdfFiles);

        // Tạo một đối tượng Progress để hiển thị tiến trình
        var progress = new Progress<int>(percent =>
        {
            Console.Clear(); // Xóa màn hình để hiển thị thanh tiến trình mới
            Console.WriteLine($"Đang quét... {percent}%");
        });

        // Đếm số trang và lấy thông tin khổ giấy song song
        await ProcessPdfFilesAsync(pdfFiles, progress);

        Console.WriteLine("Hoàn thành!");
    }

    // Hàm quét các file PDF trong thư mục và các thư mục con
    static void GetPdfFiles(string directoryPath, List<FileInfo> pdfFiles)
    {
        try
        {
            DirectoryInfo directory = new DirectoryInfo(directoryPath);
            FileInfo[] files = directory.GetFiles("*.pdf", SearchOption.AllDirectories);
            pdfFiles.AddRange(files);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Lỗi: " + ex.Message);
        }
    }

    // Hàm xử lý các file PDF và hiển thị thanh tiến trình
    static async Task ProcessPdfFilesAsync(List<FileInfo> pdfFiles, IProgress<int> progress)
    {
        int totalFiles = pdfFiles.Count;
        var tasks = pdfFiles.Select((pdfFile, index) => Task.Run(() =>
        {
            int pageCount = GetPdfPageCount(pdfFile.FullName);
            //var paperSize = GetPdfPaperSize(pdfFile.FullName);
            return new { pdfFile.DirectoryName, pdfFile.Name, PageCount = pageCount, PaperSize = paperSize };
        })).ToArray();

        // Cập nhật tiến trình sau khi mỗi file được xử lý
        for (int i = 0; i < tasks.Length; i++)
        {
            var result = await tasks[i];
            int progressPercent = (int)((i + 1) / (double)totalFiles * 100);
            progress.Report(progressPercent); // Gửi thông báo tiến trình
        }

        // Xuất kết quả ra file Excel
        ExportToExcel(await Task.WhenAll(tasks));
    }

    // Hàm đếm số trang của một file PDF
    static int GetPdfPageCount(string pdfFilePath)
    {
        try
        {
            using (PdfDocument document = PdfReader.Open(pdfFilePath, PdfDocumentOpenMode.Import))
            {
                return document.PageCount;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Lỗi khi đọc file PDF {pdfFilePath}: " + ex.Message);
            return 0; // Trả về 0 nếu có lỗi khi đọc file
        }
    }

    // Hàm xác định khổ giấy của file PDF (A4, A3, A0, v.v.)
    static string GetPdfPaperSize(string pdfFilePath)
    {
        try
        {
            using (PdfDocument document = PdfReader.Open(pdfFilePath, PdfDocumentOpenMode.Import))
            {
                var page = document.Pages[0];
                double width = page.Width;
                double height = page.Height;

                // Đơn vị của PdfSharp là điểm (point), 1 point = 0.3528 mm
                // Chuyển đổi sang mm
                width *= 0.3528;
                height *= 0.3528;

                // Kiểm tra và xác định khổ giấy
                if (width == 210 && height == 297)
                    return "A4"; // Khổ A4 (210mm x 297mm)
                else if (width == 297 && height == 420)
                    return "A3"; // Khổ A3 (297mm x 420mm)
                else if (width == 841 && height == 1189)
                    return "A0"; // Khổ A0 (841mm x 1189mm)
                else
                    return "Khổ giấy khác"; // Nếu không phải A4, A3, A0
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Lỗi khi đọc kích thước khổ giấy của file PDF {pdfFilePath}: " + ex.Message);
            return "Không xác định";
        }
    }

    // Hàm xuất kết quả ra file Excel
    static void ExportToExcel(dynamic[] results)
    {
        // Lấy đường dẫn thư mục chứa file EXE hiện tại
        string exeDirectory = AppDomain.CurrentDomain.BaseDirectory;

        // Tạo đường dẫn đến file Excel trong thư mục cùng cấp với file EXE
        string outputFilePath = Path.Combine(exeDirectory, "Result.xlsx");

        // Tạo workbook mới và worksheet
        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.Worksheets.Add("Results");

            // Đặt tiêu đề cho các cột
            worksheet.Cell(1, 1).Value = "Đường dẫn";
            worksheet.Cell(1, 2).Value = "Tên file";
            worksheet.Cell(1, 3).Value = "Số trang";
            //worksheet.Cell(1, 4).Value = "Khổ giấy";

            // Điền dữ liệu vào các ô
            for (int i = 0; i < results.Length; i++)
            {
                worksheet.Cell(i + 2, 1).Value = results[i].DirectoryName;
                worksheet.Cell(i + 2, 2).Value = results[i].Name;
                worksheet.Cell(i + 2, 3).Value = results[i].PageCount;
                //worksheet.Cell(i + 2, 4).Value = results[i].PaperSize;
            }

            // Lưu kết quả vào file Excel
            workbook.SaveAs(outputFilePath);
        }

        Console.WriteLine($"Kết quả đã được lưu vào: {outputFilePath}");
    }
}
