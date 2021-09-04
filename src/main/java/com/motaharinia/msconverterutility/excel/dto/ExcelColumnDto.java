/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.motaharinia.msconverterutility.excel.dto;


import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.io.Serializable;
import java.util.HashMap;

/**
 * @author https://github.com/motaharinia<br>
 * کلاس مدل تنظمیات ستونهای اکسل
 */

@AllArgsConstructor
@NoArgsConstructor
@Data
public class ExcelColumnDto implements Serializable {

    /**
     * فرمت کننده ستون
     * مثلا میخواهیم برای مقادیر صفر در ستون کلمه خیر بیاریم و برای مقادیر یک در ستون کلمه بلی بیاریم
     */
    private HashMap<Object, Object> formatterMap;
    /**
     * تنظیمات ظاهری ستون
     */
    private ExcelStyleDto style;

}
