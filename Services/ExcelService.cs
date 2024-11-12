using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using ReadingExcel.Models;
using System.Data;
using System.Reflection;

namespace ReadingExcel.Services
{
    public class ExcelService
    {
        public static void ExportJSON(){
            string filePath = @"C:\Users\PC\source\repos\ReadingExcel\ExcelFiles\Demo.xlsx"; // Đường dẫn tới file Excel

            // Danh sách để chứa các đối tượng CustomerInfo
            List<CustomerInfo> CustomerInfos = new List<CustomerInfo>();

            // Mở và đọc file Excel
            using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                // Đọc workbook từ file Excel
                IWorkbook workbook = new XSSFWorkbook(file);
                ISheet sheet = workbook.GetSheetAt(0);  // Chọn sheet đầu tiên

                // Lấy dòng tiêu đề (dòng đầu tiên của sheet)
                IRow headerRow = sheet.GetRow(0);

                // Duyệt qua từng dòng trong sheet, bắt đầu từ dòng thứ 1 (bỏ qua dòng tiêu đề)
                for (int rowIndex = 1; rowIndex <= sheet.LastRowNum; rowIndex++)
                {
                    IRow row = sheet.GetRow(rowIndex);
                    if (row == null) continue;  // Kiểm tra dòng có dữ liệu không

                    // Tạo đối tượng CustomerInfo mới
                    CustomerInfo CustomerInfo = new CustomerInfo();

                    // Duyệt qua từng cột trong dòng
                    for (int colIndex = 0; colIndex < headerRow.LastCellNum; colIndex++)
                    {
                        ICell cell = row.GetCell(colIndex);
                        string columnName = headerRow.GetCell(colIndex).StringCellValue.Trim();  // Tên cột từ dòng tiêu đề

                        // Lấy thuộc tính có attribute [ExcelColumn] tương ứng với tên cột
                        PropertyInfo property = typeof(CustomerInfo).GetProperties()
                            .FirstOrDefault(p => p.GetCustomAttribute<ExcelColumnAttribute>()?.ColumnName == columnName);

                        if (property != null && cell != null)
                        {
                            // Gán giá trị từ cell vào thuộc tính
                            if (property.PropertyType == typeof(string))
                            {
                                property.SetValue(CustomerInfo, cell.ToString());
                            }
                            else if (property.PropertyType == typeof(DateTime?) && cell.CellType == CellType.Numeric)
                            {
                                property.SetValue(CustomerInfo, cell.DateCellValue);
                            }
                            else if (property.PropertyType == typeof(decimal) && cell.CellType == CellType.Numeric)
                            {
                                property.SetValue(CustomerInfo, (decimal)cell.NumericCellValue);
                            }
                        }
                        
                    }
                    // Kiểm tra nếu Mã khách hàng (CustomerCode) là null hoặc rỗng thì dừng lại
                    if (string.IsNullOrEmpty(CustomerInfo.CustomerCode))
                    {
                        break;  // Dừng vòng lặp nếu không có mã khách hàng
                    }
                    // Thêm CustomerInfo vào danh sách
                    CustomerInfos.Add(CustomerInfo);
                }
            }

            // In kết quả
            foreach (var CustomerInfo in CustomerInfos)
            {
                Console.WriteLine($"CustomerInfo Code: {CustomerInfo.CustomerCode}, Name: {CustomerInfo.CustomerName}");
            }
        }
    }
}
