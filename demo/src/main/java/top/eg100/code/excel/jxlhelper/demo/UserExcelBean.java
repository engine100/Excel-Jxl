package top.eg100.code.excel.jxlhelper.demo;


import jxl.format.Alignment;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.Colour;
import jxl.format.UnderlineStyle;
import jxl.format.VerticalAlignment;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WriteException;
import top.eg100.code.excel.jxlhelper.annotations.ExcelContent;
import top.eg100.code.excel.jxlhelper.annotations.ExcelContentCellFormat;
import top.eg100.code.excel.jxlhelper.annotations.ExcelSheet;
import top.eg100.code.excel.jxlhelper.annotations.ExcelTitleCellFormat;

/**
 * 用户表，作为用户的导出Excel的中间格式化实体，所有字段都为 String
 */
@ExcelSheet(sheetName = "用户表")
public class UserExcelBean {

	@ExcelContent(titleName = "姓名",index = 3)
	private String Name;

	@ExcelContent(titleName = "性别",index = 2)
	private String Sex;

	@ExcelContent(titleName = "地址",index = 4)
	private String Address;

	@ExcelContent(titleName = "电话",index = 5)
	private String Mobile;

	@ExcelContent(titleName = "其他",index = 1)
	private String Other;

	@ExcelContent(titleName = "备注",index = 0)
	private String Memo;

	@ExcelTitleCellFormat(titleName = "姓名")
	private static WritableCellFormat getTitleFormat() {
		WritableCellFormat format = new WritableCellFormat();
		try {
			// 单元格格式
			// 背景颜色
			// format.setBackground(Colour.PINK);
			// 边框线
			format.setBorder(Border.BOTTOM, BorderLineStyle.THIN, Colour.RED);
			// 设置文字居中对齐方式;
			format.setAlignment(Alignment.CENTRE);
			// 设置垂直居中;
			format.setVerticalAlignment(VerticalAlignment.CENTRE);
			// 设置自动换行
			format.setWrap(false);

			// 字体格式
			WritableFont font = new WritableFont(WritableFont.ARIAL);
			// 字体颜色
			font.setColour(Colour.BLUE2);
			// 字体加粗
			font.setBoldStyle(WritableFont.BOLD);
			// 字体加下划线
			font.setUnderlineStyle(UnderlineStyle.SINGLE);
			// 字体大小
			font.setPointSize(20);
			format.setFont(font);

		} catch (WriteException e) {
			e.printStackTrace();
		}
		return format;
	}

	private static int f1flag = 0;
	private static int f2flag = 0;
	private static int f3flag = 0;
	private static int f4flag = 0;
	private static int f5flag = 0;
	private static int f6flag = 0;

	@ExcelContentCellFormat(titleName = "姓名")
	private WritableCellFormat f1() {
		WritableCellFormat format = null;
		try {
			format = new WritableCellFormat();
			if ((f1flag & 1) != 0) {
				format.setBackground(Colour.GRAY_25);
			}

			if (Name.contains("4")) {
				format.setBackground(Colour.RED);
			}

			f1flag++;
		} catch (WriteException e) {
			e.printStackTrace();
		}
		return format;
	}

	@ExcelContentCellFormat(titleName = "性别")
	private WritableCellFormat f2() {
		WritableCellFormat format = null;
		try {
			format = new WritableCellFormat();
			if ((f2flag & 1) != 0) {
				format.setBackground(Colour.GRAY_25);
			}
			f2flag++;
		} catch (WriteException e) {
			e.printStackTrace();
		}
		return format;
	}

	@ExcelContentCellFormat(titleName = "地址")
	private WritableCellFormat f3() {
		WritableCellFormat format = null;
		try {
			format = new WritableCellFormat();
			if ((f3flag & 1) != 0) {
				format.setBackground(Colour.GRAY_25);
			}
			f3flag++;
		} catch (WriteException e) {
			e.printStackTrace();
		}
		return format;
	}

	@ExcelContentCellFormat(titleName = "电话")
	private WritableCellFormat f4() {
		WritableCellFormat format = null;
		try {
			format = new WritableCellFormat();
			if ((f4flag & 1) != 0) {
				format.setBackground(Colour.GRAY_25);
			}
			f4flag++;
		} catch (WriteException e) {
			e.printStackTrace();
		}
		return format;
	}

	@ExcelContentCellFormat(titleName = "其他")
	private WritableCellFormat f5() {
		WritableCellFormat format = null;
		try {
			format = new WritableCellFormat();
			if ((f5flag & 1) != 0) {
				format.setBackground(Colour.GRAY_25);
			}
			f5flag++;
		} catch (WriteException e) {
			e.printStackTrace();
		}
		return format;
	}

	@ExcelContentCellFormat(titleName = "备注")
	private WritableCellFormat f6() {
		WritableCellFormat format = null;
		try {
			format = new WritableCellFormat();
			if ((f6flag & 1) != 0) {
				format.setBackground(Colour.GRAY_25);
			}
			f6flag++;
		} catch (WriteException e) {
			e.printStackTrace();
		}
		return format;
	}

	public UserExcelBean() {

	}

	public String getName() {
		return Name;
	}

	public void setName(String name) {
		Name = name;
	}

	public String getSex() {
		return Sex;
	}

	public void setSex(String sex) {
		Sex = sex;
	}

	public String getAddress() {
		return Address;
	}

	public void setAddress(String address) {
		Address = address;
	}

	public String getMobile() {
		return Mobile;
	}

	public void setMobile(String mobile) {
		Mobile = mobile;
	}

	public String getOther() {
		return Other;
	}

	public void setOther(String other) {
		Other = other;
	}

	public String getMemo() {
		return Memo;
	}

	public void setMemo(String memo) {
		Memo = memo;
	}

}
