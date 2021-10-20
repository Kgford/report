# This demo represents pygal Charts
#http://www.pygal.org/en/stable/documentation/types/histogram.html
def LineHistogramView(request):

    line_chart.title = 'Hours spent in each activity during the specified e_chart = pygal.Line()
    line_chart.title = 'Hours spent in each activity during the specified 
    period ({} to {})'.format(input_start,input_end)
    line_chart.x_labels = map(lambda d: d.strftime('%Y-%m-%d'), 
    list(df.index.values))
    for activity in act_hours:
        line_chart.add(activity, df[activity])

    static_path = 'images/charts/line_chart_{}.svg'.format(user.id)
    line_chart.render_to_file('static/' + static_path)

    activities['static_path'] = static_path

    return render(request, 'gCalData/gCalData_result.html', {'script_result': script_result})




def BarHistogramView(request):


    hist = pygal.Histogram()
    hist.add('Wide bars', [(5, 0, 10), (4, 5, 13), (2, 0, 15)])
    hist.add('Narrow bars',  [(10, 1, 2), (12, 4, 4.5), (8, 11, 13)])
    hist.render()

    return render(request, 'gCalData/gCalData_result.html', {'script_result': script_result})
