using System.Collections.Generic;

namespace MiniSoftware
{
    /// <summary>
    /// Rich-text segment — kết quả convert từ BBCode (b/i/u/s/color/list). MiniWord thay
    /// placeholder bằng chuỗi RunProperties + Text tương ứng (xem AddRichText).
    /// </summary>
    public class MiniWordRichText
    {
        public string Text { get; set; }
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Underline { get; set; }
        public bool Strike { get; set; }
        /// <summary>Màu chữ dạng "#RRGGBB" hoặc "RRGGBB".</summary>
        public string Color { get; set; }
        /// <summary>Thêm ngắt dòng (Break) TRƯỚC segment — dùng cho list item và \n trong text.</summary>
        public bool NewLineBefore { get; set; }
    }
}
