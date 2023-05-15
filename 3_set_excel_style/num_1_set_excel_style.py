"""
    设置excel样式:
        1. 调整列宽:
            语法: worksheet_object.column_dimensions['列名'].width = 列宽
        2. 定义单元格样式:
            a. 边框样式 ==> Cell.Border:
                边框样式由类 Border 对象来定义; 线条由 Side 对象定义
            b. 颜色填充 ==> Cell.fill:
                颜色填充由类 PatternFill 对象定义
            c. 对齐方式 Cell.alignment:
                对齐方式由类 Alignment 对象定义
"""
