from github import Github, Auth
from dotenv import load_dotenv
import os # Used to access the .env file
from enum import Enum

import xlsxwriter

# Custom wrapper class for the Github Repository
from GithubData import GithubData

ORGANAZATION_NAME = "UPRM-CIIC4010-F23"
PROJECT_PREFIX = "pa0"
FILE_NAME = PROJECT_PREFIX + "-Distribution.xlsx"

class GradingStatus(Enum):
    NOT_GRADED = "Not Graded"
    GRADED_POOR = "Graded (Poor)"
    GRADED_LATE = "Graded (Late)"
    GRADED = "Graded"
    GRADED_EXCEPTIONAL = "Graded (Exceptional)"

def open_workbook():
    workbook = xlsxwriter.Workbook(FILE_NAME)
    worksheet = workbook.add_worksheet()

    # Define the header format
    header_cell = workbook.add_format()
    header_cell.set_bold(True)
    header_cell.set_align('center')
    header_cell.set_bg_color('#D3D3D3') # Light Gray

    # Create the header row
    worksheet.write('A1', 'Team Name', header_cell)
    worksheet.write('B1', 'Github Link', header_cell)
    worksheet.write('C1', 'Member 1', header_cell)
    worksheet.write('D1', 'Member 2', header_cell)
    worksheet.write('E1', 'TA', header_cell)
    worksheet.write('F1', 'Grading Status', header_cell)
    worksheet.write('G1', 'Comment', header_cell)

    return (workbook, worksheet)

def get_token():
    # Load the .env file
    load_dotenv()
    return os.getenv("GITHUB_TOKEN")

def get_repositories():
    g = login(get_token())

    organization = g.get_organization(ORGANAZATION_NAME)
    print("Entered Organization: " + organization.login)

    return [
        GithubData(repo) for repo in organization.get_repos() 
        if repo.name.lower().startswith(PROJECT_PREFIX) 
        and repo.name.lower() != PROJECT_PREFIX
    ]

def login(token):
    auth = Auth.Token(token)
    g = Github(auth=auth)
    print("Successfully Logged in as: " + g.get_user().login)
    return g

def main(): 
    repositories = get_repositories()
    workbook, worksheet = open_workbook()

    # Create the conditional formatting formats
    red_format = workbook.add_format({'bg_color': '#F8CECC'})
    green_format = workbook.add_format({'bg_color': '#C6EFCE'})
    blue_format = workbook.add_format({'bg_color': '#BDD7EE'})

    status_to_format = {
        # No grading by default is blank
        GradingStatus.GRADED_POOR: red_format,
        GradingStatus.GRADED_LATE: green_format,
        GradingStatus.GRADED: green_format,
        GradingStatus.GRADED_EXCEPTIONAL: blue_format
    }

    repositories = sorted(repositories, key=lambda repo: repo.get_member_count(), reverse=True)

    for i, data in enumerate(repositories):
        # Grab the team name and member count
        worksheet.write('A%s' % (i+2), data.get_team().name)
        worksheet.write('B%s' % (i+2), data.get_repository().html_url)

        # Write a comment if the team does not have 2 members
        if data.get_member_count() != 2:
            worksheet.write('G%s' % (i+2), "Member Count: %s" % data.get_member_count(), workbook.add_format({'bg_color': '#E6B8B7'}))

        # By default, the grading status is not graded
        worksheet.write('F%s' % (i+2), GradingStatus.NOT_GRADED.value)

        # Create a dropdown for the grading status
        worksheet.data_validation(
            'F%s' % (i+2),
            {'validate': 'list', 'source': [status.value for status in GradingStatus]}
        )


        # Apply Conditional Formatting to the entire row, based on the value of column F
        for status, format in status_to_format.items():
            worksheet.conditional_format('A%s:G%s' % ((i+2), (i+2)),
                                         {
                                             'type': 'formula', 
                                             'criteria': '=$F$%d="%s"' % ((i+2), status.value), 
                                             'format': format
                                        })



    worksheet.autofit()
    workbook.close()

    print("Successfully created %s" % FILE_NAME)

if __name__ == "__main__":
    main()