using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadingExcel.Models
{
    public struct CustomerGroup
    {
        /// <summary>
        /// Diamond
        /// </summary>
        public const string Diamond = "Diamond";
        /// <summary>
        /// Diamond
        /// </summary>
        public const string Gold = "Gold";
    }
    public class CustomerInfo
    {
        /// <summary>
        /// Loại khách
        /// </summary>
        [ExcelColumn("Loại khách")]
        public string CustomerType { get; set; }

        /// <summary>
        /// Chi nhánh tạo
        /// </summary>
        [ExcelColumn("Chi nhánh tạo")]
        public string BranchCreated { get; set; }

        /// <summary>
        /// Mã khách hàng
        /// </summary>
        [ExcelColumn("Mã khách hàng")]
        public string CustomerCode { get; set; }

        /// <summary>
        /// Tên khách hàng
        /// </summary>
        [ExcelColumn("Tên khách hàng")]
        public string CustomerName { get; set; }

        /// <summary>
        /// Điện thoại
        /// </summary>
        [ExcelColumn("Điện thoại")]
        public string PhoneNumber { get; set; }

        /// <summary>
        /// Địa chỉ
        /// </summary>
        [ExcelColumn("Địa chỉ")]
        public string Address { get; set; }

        /// <summary>
        /// Khu vực giao hàng
        /// </summary>
        [ExcelColumn("Khu vực giao hàng")]
        public string DeliveryArea { get; set; }

        /// <summary>
        /// Phường/Xã
        /// </summary>
        [ExcelColumn("Phường/Xã")]
        public string WardOrCommune { get; set; }

        /// <summary>
        /// Công ty
        /// </summary>
        [ExcelColumn("Công ty")]
        public string Company { get; set; }

        /// <summary>
        /// Mã số thuế
        /// </summary>
        [ExcelColumn("Mã số thuế")]
        public string TaxCode { get; set; }

        /// <summary>
        /// Số CMND/CCCD
        /// </summary>
        [ExcelColumn("Số CMND/CCCD")]
        public string IdentificationNumber { get; set; }

        /// <summary>
        /// Ngày sinh
        /// </summary>
        [ExcelColumn("Ngày sinh")]
        public DateTime? DateOfBirth { get; set; }

        /// <summary>
        /// Giới tính
        /// </summary>
        [ExcelColumn("Giới tính")]
        public string Gender { get; set; }

        /// <summary>
        /// Email
        /// </summary>
        [ExcelColumn("Email")]
        public string Email { get; set; }

        /// <summary>
        /// Facebook
        /// </summary>
        [ExcelColumn("Facebook")]
        public string Facebook { get; set; }

        /// <summary>
        /// Nhóm khách hàng
        /// </summary>
        [ExcelColumn("Nhóm khách hàng")]
        public string CustomerGroup { get; set; }

        /// <summary>
        /// Ghi chú
        /// </summary>
        [ExcelColumn("Ghi chú")]
        public string Notes { get; set; }

        /// <summary>
        /// Điểm hiện tại
        /// </summary>
        [ExcelColumn("Điểm hiện tại")]
        public decimal CurrentPoints { get; set; }

        /// <summary>
        /// Tổng điểm
        /// </summary>
        [ExcelColumn("Tổng điểm")]
        public decimal TotalPoints { get; set; }

        /// <summary>
        /// Người tạo
        /// </summary>
        [ExcelColumn("Người tạo")]
        public string CreatedBy { get; set; }

        /// <summary>
        /// Ngày tạo
        /// </summary>
        [ExcelColumn("Ngày tạo")]
        public DateTime CreatedDate { get; set; }

        /// <summary>
        /// Ngày giao dịch cuối
        /// </summary>
        [ExcelColumn("Ngày giao dịch cuối")]
        public DateTime LastTransactionDate { get; set; }

        /// <summary>
        /// Nợ cần thu hiện tại
        /// </summary>
        [ExcelColumn("Nợ cần thu hiện tại")]
        public decimal CurrentDebt { get; set; }

        /// <summary>
        /// Tổng bán
        /// </summary>
        [ExcelColumn("Tổng bán")]
        public decimal TotalSales { get; set; }

        /// <summary>
        /// Tổng bán trừ trả hàng
        /// </summary>
        [ExcelColumn("Tổng bán trừ trả hàng")]
        public decimal NetTotalSales { get; set; }

        /// <summary>
        /// Trạng thái
        /// </summary>
        [ExcelColumn("Trạng thái")]
        public string Status { get; set; }
    }
}
