from __future__ import division
from pygal.graph.graph import Graph
import pygal
# Import Bar and Line chart
from pygal.graph.histogram import Histogram
from pygal.graph.bar import Bar
from pygal.graph.line import Line
from pygal.graph.xy import XY


#https://github.com/Kozea/pygal/issues/516
class LineHist(Histogram,Line):
    
    def __init__(self, config=None, **kwargs):
        super(LineHist, self).__init__(config=config, **kwargs)
        self.y_title_secondary = kwargs.get('y_title_secondary')
        self.plotas = kwargs.get('plotas', 'line')
       
           
    # All the series are bar except the last which is line
    # It was my use case but you can change it depending on your need
        
        def _make_y_title(self):
            super(LineBar, self)._make_y_title()
            
            # Add secondary title
            if self.y_title_secondary:
                yc = self.margin_box.top + self.view.height / 2
                xc = self.width - 10
                text2 = self.svg.node(
                    self.nodes['title'], 'text', class_='title',
                    x=xc,
                    y=yc
                )
                text2.attrib['transform'] = "rotate(%d %f %f)" % (
                    -90, xc, yc)
                text2.text = self.y_title_secondary
        
        def _plot(self):
            for i, serie in enumerate(self.series, 1):
                plottype = self.plotas

                raw_series_params = self.svg.graph.raw_series[serie.index][1]
                if 'plotas' in raw_series_params:
                    plottype = raw_series_params['plotas']
                    
                if plottype == 'bar':
                    self.histogram(serie)
                elif plottype == 'line':
                    self.line(serie)
                else:
                    raise ValueError('Unknown plottype for %s: %s'%(serie.title, plottype))
        
            for i, serie in enumerate(self.secondary_series, 1):
                plottype = self.plotas

                raw_series_params = self.svg.graph.raw_series[serie.index][1]
                if 'plotas' in raw_series_params:
                    plottype = raw_series_params['plotas']
                    
                if plottype == 'bar':
                    self.histogram(serie)
                elif plottype == 'line':
                    self.line(serie)
                else:
                    raise ValueError('Unknown plottype for %s: %s'%(serie.title, plottype))
            
               
            
            
class LineBar(pygal.Line, pygal.Bar):
    def __init__(self, config=None, **kwargs):
        super(LineBar, self).__init__(config=config, **kwargs)
        self.y_title_secondary = kwargs.get('y_title_secondary')
        self.plotas = kwargs.get('plotas', 'line')

    def _make_y_title(self):
        super(LineBar, self)._make_y_title()
        
        # Add secondary title
        if self.y_title_secondary:
            yc = self.margin_box.top + self.view.height / 2
            xc = self.width - 10
            text2 = self.svg.node(
                self.nodes['title'], 'text', class_='title',
                x=xc,
                y=yc
            )
            text2.attrib['transform'] = "rotate(%d %f %f)" % (
                -90, xc, yc)
            text2.text = self.y_title_secondary

    def _plot(self):
        for i, serie in enumerate(self.series, 1):
            plottype = self.plotas

            raw_series_params = self.svg.graph.raw_series[serie.index][1]
            if 'plotas' in raw_series_params:
                plottype = raw_series_params['plotas']
                
            if plottype == 'bar':
                self.bar(serie)
            elif plottype == 'line':
                self.line(serie)
            else:
                raise ValueError('Unknown plottype for %s: %s'%(serie.title, plottype))

        for i, serie in enumerate(self.secondary_series, 1):
            plottype = self.plotas

            raw_series_params = self.svg.graph.raw_series[serie.index][1]
            if 'plotas' in raw_series_params:
                plottype = raw_series_params['plotas']

            if plottype == 'bar':
                self.bar(serie, True)
            elif plottype == 'line':
                self.line(serie, True)
            else:
                raise ValueError('Unknown plottype for %s: %s'%(serie.title, plottype))
                
'''~~~~~~~~~~~~~~~code for LineBar~~~~~~~~~~~~~~~
# plot a dashboard for month time intervals:
# 1) number of open (backlog) tickets at that time
# 2) number of new tickets in interval
# 3) number of tickets closed in interval
# 4) in seperate graph look at turnaround time for tickets resolved
#    between X and X+T
# 5) in seperate graph look at first contact time for tickets resolved
#    between X and X+T

data = [('Apr 10', 20, 30,  5, 20, 3.2),
        ('May 10', 45, 33,  5, 20, 1.7),
        ('Jun 10', 73, 30, 20, 10, 2.5),
        ('Jul 10', 83, 12, 37, 28, 3.7),
        ('Aug 10', 58, 27, 23, 18, 1.9),
        ('Sep 10', 62, 10, 23, 11, 3.8),
        ('Oct 10', 49, 17, 29, 31, 3.6),
        ('Nov 10', 31, 27, 23, 13, 1.7),
        ('Dec 10', 35, 17, 32, 44, 0.9),
        ('Jan 11', 20, 30,  5, 24, 1.7),
        ('Feb 11', 45, 33,  5, 20, 8.6),
        ('Mar 11', 73, 30, 20, 10, 3.7),
        ('Apr 11', 83, 12, 37, 28, 2.1),]

config = pygal.Config()


# Would prefer legend_at_bottom = False. So legend is next to correct
# axis for plot. However this pushes the y_title_secondary away from
# the axis.  To compensate, set legend_at_bottom_columns to 3 so first
# row is left axis and second row is right axis. With second axis plot
# showing printed values, this should reduce confusion.

# Make range and secondary range integer mutiples so I end up with
# integer values on both axes.

style=pygal.style.DefaultStyle(value_font_size=8)

chart = LineBar(config,
                width=600,
                height=300,
                title="Tracker Dashboard",
                x_title="Month",
                y_title='Count',
                y_title_secondary='Days',
                legend_at_bottom=True,
                legend_at_bottom_columns=3,
                legend_box_size=10,
                range = (0,90), # Without this the bars start below the bottom axis
                secondary_range=(0,45),
                x_label_rotation=45,
                print_values=True,
                print_values_position='top',
                style=style,
                )

chart.x_labels = [ x[0] for x in data ]
chart.x_labels.append("") # without this the final bars overlap the secondary axis

chart.add("backlog",[ x[1] for x in data] , plotas='bar')
chart.add("new",[ x[2] for x in data] , plotas='bar')
chart.add("resolved", [ x[3] for x in data] , plotas='bar')
chart.add("turnaround time", [ x[4] for x in data] , plotas='line', secondary=True)

chart.render_to_file("plotdash_pygal.svg", pretty_print=True)
'''
        