import ezElead, getpass, openpyxl, time, pickle, re, os, json

home_dir = str(os.getcwd())
download_dir = home_dir + "\\reports"

if not os.path.exists(download_dir):
    os.makedirs(download_dir)

datestamp = time.strftime("%m-%d-%y")


def break_up(csl):  # Splits strings on commas, stripping any spaces. Used for report num entry.
    reports = []
    for each in csl.split(","):
        reports.append(each.strip())
    for each in reports:
        if len(each) < 1:
            reports.remove(each)

    return reports


# class ReportForm(object):
#     """Contains all the information about the form for each report. i.e., the number of the report of the form, the
#     parameters that are changeable for the report, and the final form to be submitted on report request."""
#     def __init__(self, report_num, session):
#         self.session = session
#         self.report_num = report_num
#         self.params = self.session.get_report_params(self.report_num)
#         self.params_readable = self.init_params()
#         self.initial_form = self.session.get_report_form(report_num)
#
#         self.modify_form_wrapper()
#
#     def init_params(self):  # Tears apart the dictionaries built by ezElead into easily iterable lists inside a list.
#         parameters = []     # First index being the name of the of the param, and every other being the options
#
#         for param_value in self.params:
#             this_param = []
#             for key in self.params[param_value].keys():
#                 this_param.append(key)
#             for option_value in self.params[param_value]:
#                 for key in self.params[param_value][option_value]:
#                     this_param.append(self.params[param_value][option_value][key])
#             parameters.append(this_param)
#
#         return parameters
#
#     def modify_form_wrapper(self):  # iterates through each of the parameters and asks if the user wants to modify it.
#         if type(self.params_readable) is not list:
#             raise TypeError("Parameters should be a list before attempting to modify. Why do you see this?")
#
#         for param in self.params_readable:
#             if input("Would you like to modify {} [Y/N]".format(param[0])) == "Y":
#                 for option in param[1:]:
#                     print("{}: {}".format(param.index(option), option))
#                 self.modify_parameter(param)
#
#     def modify_parameter(self, parameter):  # Lists the options for each param and asks the user to choose one. The
#         if type(parameter) is not list:     # key on the RB form (final_form) is then changed to this value.
#             raise TypeError("Invalid parameter for modification")
#         param_selection_index = 1
#         try:
#             param_selection_index = int(input("Select an option by number: "))
#         except KeyError:
#             print("Invalid selection number.")
#             self.modify_parameter(parameter)
#
#         param_to_change = parameter[0]
#         chosen_option = parameter[param_selection_index]
#
#         param_num = None
#         option_num = None
#
#         for param_values in self.params.values():
#             if param_to_change == list(param_values.keys())[0]:
#                 param_num = list(self.params.keys())[list(self.params.values()).index(param_values)]
#                 for option_values in param_values.values():
#                     for value in option_values.values():
#                         if chosen_option == value:
#                             option_num = list(option_values.keys())[list(option_values.values()).index(chosen_option)]
#
#         self.final_form[param_num] = option_num


class Report(object):
    """Object for an individual report. Contains the form for that report, as well as the report number. Methods
    for downloading the report, etc."""
    def __init__(self, report_num, session):
        self.session = session
        self.report_num = report_num
        self.form = None
        self.params_raw = session.get_report_params(self.report_num)
        self.params_readable = self.init_params()
        self.selected_params = {}

        if type(report_num) is not str:
            raise TypeError("Report Number is not a string.")

    def init_params(self):
        parameters = []

        for param_value in self.params_raw:
            this_param = []
            for key in self.params_raw[param_value].keys():
                this_param.append(key)
            for option_value in self.params_raw[param_value]:
                for key in self.params_raw[param_value][option_value]:
                    this_param.append(self.params_raw[param_value][option_value][key])
            parameters.append(this_param)

        return parameters

    def build_form(self, **kwargs):  # Builds the form if necessary. i.e., not pulling from a favorite.
        self.form = self.session.get_report_form(self.report_num)

        if "favorite" in kwargs:
            for key in kwargs["favorite"]:
                self.form[key] = kwargs["favorite"][key]

        else:
            for param in self.params_readable:
                if input("Would you like to modify {} [Y/N]".format(param[0])) == "Y":
                    for option in param[1:]:
                        print("{}: {}".format(param.index(option), option))
                    self.modify_parameter(param)

    def modify_parameter(self, parameter):  # Lists the options for each param and asks the user to choose one. The
        if type(parameter) is not list:     # key on the RB form (final_form) is then changed to this value.
            raise TypeError("Invalid parameter for modification")
        param_selection_index = 1
        try:
            param_selection_index = int(input("Select an option by number: "))
        except KeyError:
            print("Invalid selection number.")
            self.modify_parameter(parameter)

        param_to_change = parameter[0]
        chosen_option = parameter[param_selection_index]

        param_num = None
        option_num = None

        for param_values in self.params_raw.values():
            if param_to_change == list(param_values.keys())[0]:
                param_num = list(self.params_raw.keys())[list(self.params_raw.values()).index(param_values)]
                for option_values in param_values.values():
                    for value in option_values.values():
                        if chosen_option == value:
                            option_num = list(option_values.keys())[list(option_values.values()).index(chosen_option)]

        self.form[param_num] = option_num
        self.selected_params[param_num] = option_num

    def download(self, **kwargs): # downloads the form. See ezElead for how it parses the gvReport.
        if self.form is not None:
            report_as_list = self.session.get_report(self.report_num, report_form=self.form)
        else:
            report_as_list = self.session.get_report(self.report_num)
        active_workbook = openpyxl.Workbook()
        report_sheet = active_workbook.get_active_sheet()

        for line, row in zip(report_as_list, report_sheet.iter_rows(min_col=0, min_row=0, max_col=len(report_as_list[0])
                                                                    , max_row=len(report_as_list))):
            print(line)
            for entry, cell in zip(line, row):
                try:
                    entry = float(entry)
                except ValueError:
                    pass
                cell.value = entry

        os.chdir(home_dir + "\\reports")

        if "fav_name" in kwargs:
            active_workbook.save("{}_{}.xlsx".format(kwargs["fav_name"], datestamp))
        else:
            active_workbook.save("{}_{}.xlsx".format(self.session.reports[self.report_num], datestamp))

        wait_for = input()

        os.chdir(home_dir)


class Menu(object):
    """Basic Menu, containing options and a prompt"""
    def __init__(self):
        self.options = []
        self.prompt = ""
        self.input = None
        os.system("cls")

    def print_options(self):  # Prints the menu options, obviously.
        print("\n\n")
        for option in self.options:
            print("{}: {}".format(self.options.index(option), option))

    def get_options_input(self):  # Ugh, getter methods.
        self.input = input("Choose a selection: ")

    def print_prompt(self):  # self explanatory
        print(self.prompt)

    def read(self):  # Does everything. This just makes things a little slimmer.
        self.print_prompt()
        self.print_options()
        self.get_options_input()


class LogInMenu(Menu):
    def __init__(self):
        Menu.__init__(self)
        self.options = [
            "Log In",
            "Quit"
        ]
        self.prompt = \
            """
___________________________________________________________________________

:::::::::: ::::::::: :::::::::: :::        ::::::::::     :::     :::::::::  
:+:             :+:  :+:        :+:        :+:          :+: :+:   :+:    :+: 
+:+            +:+   +:+        +:+        +:+         +:+   +:+  +:+    +:+ 
+#++:++#      +#+    +#++:++#   +#+        +#++:++#   +#++:++#++: +#+    +:+ 
+#+          +#+     +#+        +#+        +#+        +#+     +#+ +#+    +#+ 
#+#         #+#      #+#        #+#        #+#        #+#     #+# #+#    #+# 
########## ######### ########## ########## ########## ###     ### #########
___________________________________________________________________________ᶜᶠ
         """


class CredMenu(Menu):
    def __init__(self):
        Menu.__init__(self)
        self.credentials = []
        self.get_credentials()

    def get_credentials(self):
        self.credentials.append(input("Enter username: "))
        self.credentials.append(getpass.getpass("Enter password: "))


class MainMenu(Menu):
    def __init__(self):
        Menu.__init__(self)
        self.options = [
            "Get Reports",
            "Search eLead",
            "Advanced Functions"
        ]
        self.prompt = "HOME"


class ReportMenu(Menu):
    def __init__(self, session):
        Menu.__init__(self)
        self.favorites = self.favorite_load()
        self.groups = self.group_load()
        self.prompt = "Enter report selection for download, separating multiple reports with a comma."
        self.session = session
        self.print_reports()

    @staticmethod
    def favorite_load():  # loads the favorites from a file. If they aren't there, just make it a blank dictionary.
        try:
            favorites = pickle.load(open("favorites.p", "rb"))
        except FileNotFoundError:
            favorites = {}

        return favorites

    def favorite_save(self, report):  # saves the the report number and form used as a dictionary that is the value of
                                    # a dictionary, the title of which is whatever name the user chooses.
        if input("Save this report configuration as a favorite?[Y/N]") == "Y":
            favorite = {
                "report_num": report.report_num,
                "form": report.selected_params
            }
            name = str(input("Enter a name for this favorite: "))
            self.favorites[name] = favorite
            pickle.dump(self.favorites, open("favorites.p",  "wb"))

    def print_reports(self):  # self explanatory
        for report in self.session.reports:
            print("{}{} | {}".format((5 - len(report)) * " ", report, self.session.reports[report]))
        print("\n")
        for favorite in self.favorites:
            print("FAV | {}".format(favorite))
        print("\n")
        for group in self.groups:
            print("GROUP | {}".format(group))
        print("\n")

    @staticmethod
    def group_load():
        try:
            groups = pickle.load(open("groups.p", "rb"))
        except FileNotFoundError:
            groups = {}

        return groups

    def group_save(self, reports):
        if len(reports) > 1 and type(reports) is list:
            if input("Save these reports as a group?[Y/N]") == "Y":
                name = str(input("Enter a name for this report group: "))
                self.groups[name] = reports
                pickle.dump(self.groups, open("groups.p", "wb"))

    def report_get(self, reports, **kwargs):
        for entry in reports:
            if entry in self.favorites:
                report_num = self.favorites[entry]["report_num"]
                this_report = Report(report_num, self.session)
                this_report.build_form(favorite=self.favorites[entry]["form"])
                this_report.download(fav_name=entry)
            elif entry in self.session.reports:
                this_report = Report(entry, self.session)
                if "group" not in kwargs:
                    if input("Would you like to modify the options on {}[Y/N]".format(self.session.reports[entry])) == \
                            "Y":
                        this_report.build_form()
                this_report.download()
                if "group" not in kwargs:
                    self.favorite_save(this_report)
            elif entry in self.groups:
                self.report_get(self.groups[entry], group=True)
            else:
                print("{} is not a valid report, favorite, or group.".format(entry))
                wait_for = input()
                continue
        if "group" not in kwargs:
            self.group_save(reports)


class SearchMenu(Menu):
    def __init__(self, session):
        Menu.__init__(self)
        self.prompt = "Search eLead. Try a phone number, address, email, or name."
        self.session = session
        self.results = None

    def search(self):
        self.results = self.session.search(self.input)

    def read_results(self):
        for line in self.results:
            print(line)
        wait_for = input()


def main_menu(session):
    menu = MainMenu()
    menu.read()

    if menu.input == "0":
        session.get_reports()
        report_menu = ReportMenu(session)
        report_menu.read()
        report_menu.report_get(break_up(report_menu.input))
    elif menu.input == "1":
        search_menu = SearchMenu(session)
        search_menu.read()
        search_menu.search()
        search_menu.read_results()
    elif menu.input == "2":
        return

    main_menu(session)


def core():
    log_in = LogInMenu()

    log_in.read()

    session = ezElead.ELeadSession()

    if log_in.input == "0":
        cred_menu = CredMenu()
        try:
            session.log_in(cred_menu.credentials)
        except ezElead.LoginException:
            input("Invalid Login Credentials")
            core()
    elif log_in.input == "1":
        quit()
    else:
        core()

    main_menu(session)

core()
