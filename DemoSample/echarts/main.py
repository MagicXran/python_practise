from pyecharts import options as opts
from pyecharts.charts import Bar
from pyecharts.charts import Line


# 导出图片，需要引入以下对象

def main():
    # V1 版本开始支持链式调用
    bar = (Bar().add_xaxis(["衬衫", "毛衣", "领带", "裤子", "风衣", "高跟鞋", "袜子"]).add_yaxis("商家A", [114, 55, 27, 101, 125, 27,
                                                                                          105]).add_yaxis("商家B",
                                                                                                          [57, 134, 137,
                                                                                                           129, 145, 60,
                                                                                                           49]).set_global_opts(
        title_opts=opts.TitleOpts(title="某商场销售情况")))
    bar.render()

    # 不习惯链式调用的开发者依旧可以单独调用方法
    bar = Bar()
    bar.add_xaxis(["衬衫", "毛衣", "领带", "裤子", "风衣", "高跟鞋", "袜子"])
    bar.add_yaxis("商家A", [114, 55, 27, 101, 125, 27, 105])
    bar.add_yaxis("商家B", [57, 134, 137, 129, 145, 60, 49])
    bar.set_global_opts(title_opts=opts.TitleOpts(title="某商场销售情况"))
    bar.render()


def line_bar():
    year = ["1995", "1996", "1997", "1998", "1999", "2000", "2001", "2002", "2003", "2004", "2005", "2006", "2007",
            "2008", "2009"]
    postage = [0.32, 0.32, 0.32, 0.32, 0.33, 0.33, 0.34, 0.37, 0.37, 0.37, 0.37, 0.39, 0.41, 0.42, 0.44]

    (Line().set_global_opts(tooltip_opts=opts.TooltipOpts(is_show=False), xaxis_opts=opts.AxisOpts(type_="category"),
                            yaxis_opts=opts.AxisOpts(type_="value", axistick_opts=opts.AxisTickOpts(is_show=True),
                                                     splitline_opts=opts.SplitLineOpts(is_show=True), ), ).add_xaxis(
        xaxis_data=year).add_yaxis(series_name="", y_axis=postage, symbol="emptyCircle", is_symbol_show=True,
                                   label_opts=opts.LabelOpts(is_show=True), ).render("basic_line_chart.html"))


def gen_picture():
    def bar_chart() -> Bar:
        c = (Bar().add_xaxis(["衬衫", "毛衣", "领带", "裤子", "风衣", "高跟鞋", "袜子"]).add_yaxis("商家A", [114, 55, 27, 101, 125, 27,
                                                                                            105]).add_yaxis("商家B",
                                                                                                            [57, 134,
                                                                                                             137, 129,
                                                                                                             145, 60,
                                                                                                             49]).reversal_axis().set_series_opts(
            label_opts=opts.LabelOpts(position="right")).set_global_opts(title_opts=opts.TitleOpts(title="Bar-测试渲染图片")))
        return c

    bar_chart().render("test.html")


if __name__ == '__main__':
    # main()  # line_bar()
    # print(pyecharts.__version__)
    gen_picture()
