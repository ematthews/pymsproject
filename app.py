
import logging
import tkinter as tk
import matplotlib

matplotlib.use('Agg')   # This must be loaded before pyplot!

from numpy import array as numpy_array
from datetime import datetime
from xml.etree import ElementTree as ET
from matplotlib import pyplot as plt
from dateutil.parser import parse as dateparser
from tkinter.filedialog import askopenfilename
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

class Application(tk.Frame):
    def __init__(self, master=None):
        tk.Frame.__init__(self, master)
        logging.debug('Starting application window')
        self.master = master
        self.master.title("Project Information")

        self.horizon = 10
        self.show_overdue = False
        self.show_milestones = True
        self.plot_image = True
        self.save_file = 'test.png'
        self.font_size = 'x-small'
        self.dpi = 70
        self.image_width = 1000
        self.image_height = 800
        self.ybuffer_space = 0.5
        self.annotate_rotation = 38

        self.offer_filetypes = (
            ("Microsoft Project XML Files", "*.xml"),
        )

        self.z_labels = list()
        self.fig = plt.Figure(
            figsize=(self.image_width/self.dpi, self.image_height/self.dpi),
            dpi=self.dpi,
        )

        logging.debug('Setting up application components')
        self.pack()
        self.createMenu()
        self.createWidgets()

    def createMenu(self):
        """
        Configure the drop down menus
        """
        menu = tk.Menu(self.master)
        self.master.config(menu=menu)

        file = tk.Menu(tearoff=0)
        file.add_command(label='Open', command=self.open_file)
        file.add_command(label='Exit', command=lambda:exit(0))
        menu.add_cascade(label='File', menu=file)
        logging.debug('Completed application menu configuration')

    def createWidgets(self):
        """
        Configure the GUI layout
        """
        self.options_frame = tk.Frame(self.master)
        self.canvas_frame = tk.Frame(self.master)

        self.file_name = tk.Entry(self.options_frame, text='.. filename ..')
        self.load_button = tk.Button(self.options_frame, text="Browse", command=self.open_file, width=10)
        self.save_button = tk.Button(self.options_frame, text="Save", command=self.save_image, width=5)

        self.file_name.pack(fill=tk.BOTH, expand=1, side=tk.LEFT)
        self.save_button.pack(side=tk.RIGHT)
        self.load_button.pack(side=tk.RIGHT)

        self.canvas = FigureCanvasTkAgg(self.fig, master=self.canvas_frame)
        self.canvas.show()
        self.canvas.get_tk_widget().pack(side=tk.BOTTOM, fill=tk.BOTH, expand=1)

        self.options_frame.pack(side=tk.TOP, fill=tk.BOTH)
        self.canvas_frame.pack(side=tk.BOTTOM)

    def save_image(self):
        """
        Save the output image to a file
        """
        logging.info('Saving file to {s}'.format(s=self.save_file))
        self.fig.savefig(self.save_file)

    def open_file(self):
        """
        Open the file and begin work!
        """
        fname = askopenfilename(filetypes=self.offer_filetypes)

        if fname:
            logging.info('Beginning parsing file: {f}'.format(f=fname))
            self.file_name.delete(0, tk.END)
            self.file_name.insert(0, fname)

            # ps = ProjectParser()
            self.parseFile(
                horizon=self.horizon,
                filename=fname,
                show_overdue=self.show_overdue,
                show_milestones=self.show_milestones,
                plot_image=self.plot_image,
                save_file=self.save_file,
            )
        else:
            print("No file requested")

        return fname

    def getElement(self, tree=None, tag=None, ns=None):
        """
        Retrieve a *single* element in the tree
        """
        return tree.find("{ns}{tag}".format(
            ns=ns,
            tag=tag
        ))

    def getElements(self, tree=None, tag=None, ns=None):
        """
        Retrieve a *list* of elements in the tree
        """
        return tree.findall("{ns}{tag}".format(
            ns=ns,
            tag=tag
        ))

    def labelFormatter(self, tick_value, position):
        """
        Simple matplotlib formatter to replace numbers with our labels!
        """
        value = int(tick_value)
        if value >= 0 and value < len(self.z_labels):
            return self.z_labels[value]
        else:
            return ""

    def plotImage(self, z_dates, z_data, save_file):
        """
        Perform the actual contents plotting
        """
        logging.info('Plotting image')

        ymin = 0 - float(self.ybuffer_space)
        ymax = len(self.z_labels)  # + float(self.ybuffer_space)

        logging.debug('Set ymax to {y}'.format(y=ymax))

        plt.tick_params(axis='both', which='major', labelsize=self.font_size)
        plt.xlabel('Delivery Dates', fontsize=self.font_size)   # Dates
        plt.ylabel('')   # Categories
        plt.gcf().autofmt_xdate()

        for index in range(len(self.z_labels)):
            logging.debug('Plotting {i}: {l}'.format(i=index, l=self.z_labels[index]))
            ax = self.fig.add_subplot(111)
            ax.plot(
                numpy_array(z_dates[index]),
                numpy_array([index] * len(z_dates[index])),
                marker='^',
                markersize=15,
                #color='r',  # let it auto pick for colors
                linestyle='-',
            )
            for subindex in range(len(z_dates[index])):
                this_date = z_dates[index][subindex]
                txt = z_data[index][subindex]
                ax.annotate(
                    s=txt,
                    xy=(this_date, float(index) + 0.05),
                    xycoords='data',
                    rotation=self.annotate_rotation,
                    rotation_mode='anchor',
                    size=self.font_size,
                    va='bottom',
                    ha='left',
                )

        ax.set_ylim([ymin, ymax])

        locator = matplotlib.dates.MonthLocator(bymonth=[1,4,7,10])
        self.fig.gca().yaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(self.labelFormatter))
        self.fig.gca().yaxis.set_major_locator(matplotlib.ticker.MultipleLocator(1))
        self.fig.gca().xaxis.set_major_formatter(matplotlib.dates.DateFormatter('%m/%d/%Y'))
        self.fig.gca().xaxis.set_major_locator(locator)

        self.canvas.draw()

    def parseFile(
        self, horizon, filename, show_overdue, show_milestones, plot_image, save_file, ns="{http://schemas.microsoft.com/project}"
    ):
        skip_really_old = True

        # Retrieve and parse the XML file into the etree
        logging.debug('Beginning loading of XML contents')
        tree = ET.parse(filename)
        root = tree.getroot()
        now = datetime.now()
        logging.debug('Completed loading XML contents into memory')

        # Retrieve the list of tasks from the project
        logging.debug('Retrieving all tasks from project')
        all_tasks = self.getElement(root, "Tasks", ns)

        # Get the last date saved
        logging.debug('Extracting file information')
        last_saved_text = self.getElement(root, "LastSaved", ns).text
        last_saved = dateparser(last_saved_text)
        file_age = (now - last_saved).days

        # Parse out each task into an individual item
        logging.debug('Extracting individual tasks from task list')
        tasks = self.getElements(all_tasks, "Task", ns)

        # Result sets to save messages to
        upcoming = list()
        overdue = list()
        future = list()

        # Go through each task in detail
        top_level = False
        parent_task = None
        key_milestones = list()

        z_dates = list()
        z_data = list()

        logging.debug('Beginning parsing of individual tasks')
        for task in tasks:
            # ID is the line number in Project
            task_id = self.getElement(task, "ID", ns).text

            # Not sure what type is yet, haven't seen pattern
            task_type = self.getElement(task, "Type", ns).text

            # Rollup of 0 is an actual task and 1 is an outline level
            rollup = self.getElement(task, "Rollup", ns).text

            # Is it finished?
            percent_complete = self.getElement(task, "PercentComplete", ns).text

            # Get the name of the task, if it is available
            task_name = self.getElement(task, "Name", ns)
            if task_name is not None:
                name = task_name.text
            logging.debug('Parsed task: {t}'.format(t=name))

            # Alternatively use OutlineNumber
            try:
                wbs = self.getElement(task, "WBS", ns).text
                category_number = str(wbs).split('.')[0]
                if category_number == '1':
                    logging.debug('Found a top level task!')
                    top_level = True
                else:
                    top_level = False
            except AttributeError:
                top_level = False

            # Only do stuff with NOT rollup tasks
            if int(rollup) is not 0:
                # Set the last_rollup to this name
                parent_task = name
                #parent_task = str(name).encode('utf-8')
                continue

            # Also, no point looking at tasks already done
            if int(percent_complete) is 100:
                continue


            logging.debug('Found task with further interesting details')
            task_finish = self.getElement(task, "ManualFinish", ns)
            if task_finish is not None:
                finish_text = task_finish.text

            # Calculate how many days until this is supposed to be finished
            finish = dateparser(finish_text)
            days_to_go = (finish - now).days

            # Setup the line to show person
            out = "[Line {line}] Due in {due} days ({due_date}): {name} ({parent})".format(
                line=task_id,
                due=days_to_go,
                due_date=finish,
                name=name,
                parent=parent_task,
            )

            # Add this to the top level ones if applicable
            if show_milestones and top_level:
                # Skip any over a year old
                if skip_really_old and days_to_go < -400:
                    continue
                if parent_task not in self.z_labels:
                    self.z_labels.append(parent_task)
                    z_dates.append(list())
                    z_data.append(list())

                z_index = self.z_labels.index(parent_task)
                z_dates[z_index].append(matplotlib.dates.datestr2num(finish_text))
                z_data[z_index].append(name)

                my_tuple = (
                    finish_text,
                    str(parent_task).encode('utf-8'),
                    "{parent}:{name}".format(
                        name=name,
                        parent=parent_task,
                    )
                )
                key_milestones.append(my_tuple)

            # Prepend "OVERDUE" in front of any line output for anything past due date
            if days_to_go < 0:
                overdue.append(out)

            # Only show items that are within the time frame we are looking
            elif days_to_go < horizon:
                upcoming.append(out)

            # Future items
            else:
                future.append(out)
        logging.debug('Completed parsing all task information')

        # # Finally show the output to requestor
        # if show_overdue:
        #     print("------------------------------------")
        #     print(" OVERDUE!")
        #     print("------------------------------------")
        #     print()
        #     for item in overdue:
        #         print(item)

        if plot_image:
            logging.info('Plotting image contents')
            self.plotImage(z_dates, z_data, save_file)

        # print("------------------------------------")
        # print("Upcoming Tasks, next %s days" % (horizon))
        # print("    [File saved %s day(s) ago]" % (file_age))
        # print("------------------------------------")
        # for item in upcoming:
        #     print(item)


if __name__ == '__main__':
    logging.basicConfig(level=logging.DEBUG)
    root = tk.Tk()
    app = Application(master=root)
    app.mainloop()
