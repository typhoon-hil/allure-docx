import pygal
import pygal.style


def create_piechart(data, imgfile):
    font_family = "Arial"
    font_size = 25
    red = "#FF0000"
    yellow = "#FFFF00"
    green = "#00FF00"
    grey = "#DDDDDD"

    style = pygal.style.Style()
    style.colors = (yellow, red, grey, green)

    style.background = "#FFFFFF"
    style.plot_background = "#FFFFFF"

    style.font_family=font_family
    style.label_font_family=font_family
    style.legend_font_family=font_family
    style.title_font_family=font_family
    style.value_font_family=font_family
    style.value_label_font_family=font_family

    style.font_size=font_size
    style.label_font_size=font_size
    style.legend_font_size=font_size
    style.title_font_size=font_size
    style.value_font_size=font_size
    style.value_label_font_size=font_size

    config = pygal.Config()
    config.show_legend = True
    config.human_readable = True
    config.print_values=True
    config.print_labels=True
    pie_chart = pygal.Pie(config=config, style=style, inner_radius=.4)

    for item in data:
        pie_chart.add(item, data[item])

    pie_chart.render_to_png(imgfile)

