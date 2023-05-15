"""
    excel图表设置:
        坐标轴:
            1. 值轴: 用于描述数量, 标签一般是数字
            2. 类别轴: 用于描述类别, 标签一般为类别名
        图例: 说明图中数据系列的含义
        确定图的类型:
            openpyxl.chart模块中, 每个类型的图都有对应的类, 常用的折线图和条形图对应如下:
                LineChart(折线图)类实例化: chart = LineChart()
                BarChart(条形图)类实例化: chart = BarChart()
        引用表格中的数据:
            实例化 openpyxl.chart 模块中的Reference类, 可完成表格数据的引用
                Reference类实例化:
                    Reference_object = Reference(worksheet=worksheet_object,
                    min_row = start_row, max_row = end_row, min_col = start_col, max_col = end_col)
        折线图绘制:
            LineChart_object.add_data():
                语法:
                    LineChart_object.add_data(Reference_object, from_rows=False, titles_from_data=False)
                参数from_rows:
                    True: 引用区域的每一行数据绘制为一条折线
                    False: 引用区域的每一列数据绘制为一条折线
                参数titles_from_data:
                    True: 会将引用数据的首列用作图例, 其他数据用作绘制折线
                    False: 会将引用的所有数据用作绘制折线
            设置图在工作表中的位置
                语法:
                    worksheet_object.add_chart(chart_object, anchor(chart_local))
                        参数说明:
                            chart_object ==> 图表对象, 可以是折线图, 也可以是条形图等其他图表对象
                            chart_local ==> 图表左上角所在单元格位置, 在此位置插入图表, 例: C10
            折线图信息及样式优化
                a. 修改类别轴的标签
                    LineChart_object.set_categories(Reference_object):
                        Reference_object指明要引用的数据, 可以是表头部分的单元格
                b. 添加X轴, Y轴的标题
                    X轴: LineChart_object.x_axis.title = 'x轴标题'
                    Y轴: LineChart_object.y_axis.title = 'y轴标题'
                c. 修改折线图的样式 ==> LineChart_object.style
                d. 设置折线图大小:
                    设置折线图的宽度 ==> LineChart_object.width = width_num
                    设置折线图的高度 ==> LineChart_object.height = height_num
        条形图绘制 ==> 与折线图的绘制方法完全一致, 只是要创建条形图对象 BarChart_Object 操作
"""
