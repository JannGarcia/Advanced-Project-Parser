from github import Github, Auth
from dotenv import load_dotenv
import os # Used to access the .env file
from enum import Enum

import xlsxwriter
from random import shuffle
# Custom wrapper class for the Github Repository
from GithubData import GithubData

ORGANAZATION_NAME = "UPRM-CIIC4010-F23"
PROJECT_PREFIX = "pa0"
FILE_NAME = PROJECT_PREFIX + "-Distribution.xlsx"


class ColumnName(Enum):
    TEAM_NAME = "Team Name"
    GITHUB_LINK = "Github Link"
    MEMBER_1 = "Member 1"
    MEMBER_2 = "Member 2"
    TA = "TA"
    GRADING_STATUS = "Grading Status"
    COMMENT = "Comment"

class GradingStatus(Enum):
    NOT_GRADED = "Not Graded"
    GRADED_POOR = "Graded (Poor)"
    GRADED_LATE = "Graded (Late)"
    GRADED = "Graded"
    GRADED_EXCEPTIONAL = "Graded (Exceptional)"

column_name_to_index = {
    ColumnName.TEAM_NAME: 'A',
    ColumnName.GITHUB_LINK: 'B',
    ColumnName.MEMBER_1: 'C',
    ColumnName.MEMBER_2: 'D',
    ColumnName.TA: 'E',
    ColumnName.GRADING_STATUS: 'F',
    ColumnName.COMMENT: 'G'
}

def get_cell_index(column_name, i):
    return '%s%s' % (column_name_to_index[column_name], i)

def open_workbook():
    workbook = xlsxwriter.Workbook(FILE_NAME)
    worksheet = workbook.add_worksheet()

    # Define the header format
    header_cell = workbook.add_format()
    header_cell.set_bold(True)
    header_cell.set_align('center')
    header_cell.set_bg_color('#D3D3D3') # Light Gray

    # Create the header rows
    for column_name, column_index in column_name_to_index.items():
        worksheet.write('%s1' % column_index, column_name.value, header_cell)

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

    # Find the index of the first repository with less than 2 members
    reverse_index = 0
    for i, data in enumerate(repositories):
        reverse_index = i
        if data.get_member_count() != 2:
            break

    # Shuffle from 0 to reverse_index, leaving the 1 member/0 member teams
    # at the bottom.
    # This is to avoid having the most recent submissions at the top
    # and having another Christopher situation :)
    first_half = repositories[:reverse_index]
    shuffle(first_half)
    repositories = first_half + repositories[reverse_index:]


    for i, data in enumerate(repositories):
        # Grab the team name and member count
        worksheet.write(
            get_cell_index(ColumnName.TEAM_NAME, (i+2)),
            data.get_team().name
        )

        worksheet.write(
            get_cell_index(ColumnName.GITHUB_LINK, (i+2)),
            data.get_repository().html_url
        )

        # Write a comment if the team does not have 2 members
        if data.get_member_count() != 2:
            worksheet.write(
                get_cell_index(ColumnName.COMMENT, (i+2)),
                "Member Count: %s" % data.get_member_count(),
                workbook.add_format({'bg_color': '#E6B8B7'})
            )

        # By default, the grading status is not graded
        worksheet.write(
            get_cell_index(ColumnName.GRADING_STATUS, (i+2)),
            GradingStatus.NOT_GRADED.value
        )

        # Create a dropdown for the grading status
        worksheet.data_validation(
            get_cell_index(ColumnName.GRADING_STATUS, (i+2)),
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

    # Autofit the column widths, and save the file
    worksheet.autofit()
    workbook.close()

    print("Successfully created %s" % FILE_NAME)

if __name__ == "__main__":
    main()