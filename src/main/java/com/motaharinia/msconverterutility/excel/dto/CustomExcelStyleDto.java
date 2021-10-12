package com.motaharinia.msconverterutility.excel.dto;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

import java.awt.*;
import java.io.Serializable;
/**
 * @author https://github.com/motaharinia<br>
 * کلاس تنظیمات ظاهری اکسل
 */

@AllArgsConstructor
@NoArgsConstructor
@Data
public class CustomExcelStyleDto implements Serializable {
    /**
     * جهت
     */
    private HorizontalAlignment alignment = HorizontalAlignment.CENTER;
    /**
     * نام قلم
     */
    private String fontName = "Tahoma";
    /**
     * برجسته بودن قلم
     */
    private Boolean fontIsBold = false;
    /**
     * رنگ قلم
     */
    private Color fontColor = Color.BLACK;
    /**
     * رنگ پس زمینه
     */
    private Color backgroundColor = Color.WHITE;
    /**
     * نوع جدول
     */
    private BorderStyle borderStyle = BorderStyle.THIN;
    /**
     * رنگ جدول
     */
    private Color borderColor = Color.WHITE;
    /**
     * فرمت نمایش اطلاعات
     * #,##0
     * #.###############
     * m/d/yy h:mm
     */
    private String dataFormat = "General";
}
