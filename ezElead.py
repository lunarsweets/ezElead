import robobrowser
import re
import os
from pprint import pprint


class LoginException(Exception):
    """Raised when the browser finds the string "not found" after submitting the login form. This would indicate that
    the browser was unable to reach the eLead home screen."""
    pass


class InvalidReportException(Exception):
    """Raised when the report number chosen is not in the list of available reports"""
    pass


class ELeadSession(object):

    def __init__(self):
        self.session = robobrowser.RoboBrowser(history=True, parser="lxml")
        self.reports = []

    def log_in(self, credential_list):  # Logs in to eLead with credentials given as a list. [0] is user, [1] is pass
        browser = self.session

        browser.open("https://www.eleadcrm.com/evo2/fresh/login.asp")

        form = browser.get_form(action="/evo2/fresh/login.asp")
        form["user"].value = credential_list[0]
        form["Password"].value = credential_list[1]
        browser.submit_form(form)

        form_two = browser.get_form(action="/evo2/fresh/login.asp")
        browser.submit_form(form_two)

        if "not found" in browser.find("b").string:
            raise LoginException

    def get_reports(self):  # Gets all the available reports for the given eLead session
        browser = self.session

        browser.open("https://www.eleadcrm.com/evo2/fresh/elead-v45/elead_track/Reports/ReportMenu.aspx")
        links_table = browser.find("table")

        reports_dictionary = {}
        letters = re.compile(r"[\D]")

        report_blacklist = ["47", "29", "1826"]

        for element in links_table.find_all("a"):
            try:
                number_index = element["href"].index("=") + 1
                if re.search(letters, element["href"][number_index:]) \
                        or element["href"][number_index:] in report_blacklist:
                    pass
                else:
                    reports_dictionary[element["href"][number_index:]] = element.string
            except ValueError:
                pass

        self.reports = reports_dictionary

    def get_report_params(self, report_num): # Gets the report parameters for a given report
        browser = self.session

        if report_num in self.reports:
            pass
        else:
            raise InvalidReportException

        browser.open("https://www.eleadcrm.com/evo2/fresh/elead-v45/elead_track/reports"
                     "/customReport.aspx?ID={}".format(report_num))

        for row in browser.find_all("tr", id="ctlCriteriaContainer"):
            form_params = {
                param["name"]:
                    {param["parameterlabel"]:
                         {option["value"]:
                              option.string for option in param.contents if option.string != "\n"}}
                for param in row.find_all("select")
            }

            return form_params

    def get_report_form(self, report_num):
        browser = self.session

        browser.open("https://www.eleadcrm.com/evo2/fresh/elead-v45/elead_track/reports"
                     "/customReport.aspx?ID={}".format(report_num))
        form = browser.get_form(action="./customReport.aspx?ID={}".format(report_num))

        return form

    def get_report(self, report_num, **kwargs): # submits and parses the given report with the given parameters
        browser = self.session

        if report_num in self.reports:
            pass
        else:
            raise InvalidReportException

        browser.open("https://www.eleadcrm.com/evo2/fresh/elead-v45/elead_track/reports"
                     "/customReport.aspx?ID={}".format(report_num))

        report_form = browser.get_form(action="./customReport.aspx?ID={}".format(report_num))

        if "report_form" in kwargs:
            report_form = kwargs["report_form"]

        browser.submit_form(report_form)

        report = []

        for tr in browser.find("table", id="gvReport").find_all("tr"):
            i = 0
            for each in tr.descendants:
                i += 1
            if i < 100:  # this tests to make sure it's using the correct table row, instead of one with a larger scope.
                row = []
                for th in tr.find_all("th"):
                    for content in th.contents:
                        if content.previous_sibling is not None:  # to prevent from iterating over the same th/td twice
                            continue
                        elif content.string is None: # if there isn't a string
                            if content.next_sibling is None:  # if there isn't a sibling
                                continue
                            else:
                                row.append(content.next_sibling.string)  # if there is, append that instead
                        else:
                            row.append(content.string) # if there is, append it
                for td in tr.find_all("td"):
                    for content in td.contents:
                        if content.previous_sibling is not None:
                            continue
                        elif content.string is None:
                            if content.next_sibling is None:
                                continue
                            else:
                                row.append(content.next_sibling.string)
                        else:
                            row.append(content.string)
            else:
                continue
            report.append(row)

        return report

    def search(self, search_criteria):
        browser = self.session

        browser.open("https://www.eleadcrm.com/evo2/fresh/elead-v45/elead_track/"
                     "search/searchresults.asp?Go=2&searchexternal=&q={}"
                     "&st=0&lIUID=&etitle=&lDID=&PID=&origStrDo=".format(search_criteria))

        results = []

        for line in browser.find_all("tr", "InfragisticsBorderBottom textBlack wgSubCategory-ic"):
            entry = []
            for data in line:
                entry.append(data.contents)
            results.append(entry)

        return results

# session_one = ELeadSession()
# session_one.log_in(["donnaash", "Sales3"])
# session_one.get_reports()
#
# print(session_one.get_report_params("213"))

# test wrapper for printing out each parameter and their options--used for building out user desired forms
# for param_num in params.keys():
#     for label in params[param_num].keys():
#         print(label)
#         for option in params[param_num][label]:
#             print(params[param_num][label][option])


