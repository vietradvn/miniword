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
        /// <summary>
        /// Căn lề của đoạn chứa segment: "left" | "center" | "right" | "justify".
        /// Alignment là thuộc tính paragraph chứ không phải run, nên AddRichText đặt nó lên
        /// ParagraphProperties của đoạn chứa placeholder (bug TC 14 — mẫu in mất căn giữa/căn phải).
        /// </summary>
        public string Align { get; set; }
        /// <summary>
        /// Font riêng của segment (null = giữ font của placeholder). Dùng cho ký hiệu chỉ đúng hình trong một font
        /// nhất định — vd gạch đầu dòng chuẩn của Word là U+F0B7 trong font "Symbol".
        /// </summary>
        public string FontFamily { get; set; }

        /// <summary>
        /// Segment này mở một ĐOẠN (paragraph) mới thay vì nối tiếp đoạn hiện tại. Đoạn mới chép định dạng đoạn
        /// của placeholder (font/giãn dòng/căn lề…); <see cref="Align"/>, <see cref="IndentLeftTwips"/>,
        /// <see cref="IndentHangingTwips"/> của CHÍNH segment này là định dạng riêng của đoạn đó. Dùng cho mục
        /// danh sách: mỗi mục một đoạn thụt treo thì dòng xuống hàng thẳng dưới chữ chứ không về sát lề.
        /// Không segment nào bật cờ này ⇒ giữ hành vi cũ (mọi thứ trong đoạn của placeholder).
        /// </summary>
        public bool NewParagraphBefore { get; set; }

        /// <summary>Lề trái của đoạn (twip, cộng thêm vào lề sẵn có của đoạn placeholder) — chỉ đọc khi <see cref="NewParagraphBefore"/>.</summary>
        public int? IndentLeftTwips { get; set; }

        /// <summary>
        /// Thụt treo của đoạn (twip): dòng đầu bắt đầu lùi ra <c>IndentLeft - Hanging</c>, các dòng sau ở <c>IndentLeft</c>.
        /// Kèm tab stop tại <c>IndentLeft</c> để "\t" sau ký hiệu nhảy đúng tới chỗ bắt đầu chữ.
        /// </summary>
        public int? IndentHangingTwips { get; set; }
    }
}
