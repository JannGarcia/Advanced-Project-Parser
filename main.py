from math import floor, ceil

from github import Github, Auth
from dotenv import load_dotenv
import os  # Used to access the .env file
from enum import Enum

import xlsxwriter
from random import shuffle
# Custom wrapper class for the GitHub Repository
from GithubData import GithubData

ORGANIZATION_NAME = "UPRM-CIIC4010-F24"
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
    GRADED_POOR = "Graded (<=50%)"
    GRADED_LATE = "Graded (Late)"
    GRADED = "Graded"
    GRADED_EXCEPTIONAL = "Graded (>=100%)"


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
    header_cell.set_bg_color('#D3D3D3')  # Light Gray

    # Create the header rows
    for column_name, column_index in column_name_to_index.items():
        worksheet.write('%s1' % column_index, column_name.value, header_cell)

    return workbook, worksheet


def get_token():
    # Load the .env file
    load_dotenv()
    token = os.getenv("GITHUB_TOKEN")

    if not isinstance(token, str):
        print(f"Token value expected string, got {type(token)}")
        print(f"Perhaps you have not loaded your env file?")
        exit(400)

    return token


def get_repositories():
    g = login(get_token())

    organization = g.get_organization(ORGANIZATION_NAME)
    print("Entered Organization: " + organization.login)
    repos = [
        GithubData(repo) for repo in organization.get_repos()
        if repo.name.lower().startswith(PROJECT_PREFIX) and repo.name.lower() != PROJECT_PREFIX
    ]

    if not repos:
        raise ValueError(
            "No repositories found for this project.\n"
            "Check your token's permissions to allow access."
        )

    return repos


def login(token):
    auth = Auth.Token(token)
    g = Github(auth=auth)
    print("Successfully Logged in as: " + g.get_user().login)
    return g


def shuffle_until_no_two_members(repos):
    repositories = sorted(repos, key=lambda repo: repo.get_member_count(), reverse=True)

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
    return repositories


def main():
    repositories = get_repositories()
    workbook, worksheet = open_workbook()

    # Create the conditional formatting
    red_format = workbook.add_format({'bg_color': '#F8CECC'})
    green_format = workbook.add_format({'bg_color': '#C6EFCE'})
    blue_format = workbook.add_format({'bg_color': '#BDD7EE'})
    yellow_format = workbook.add_format({'bg_color': '#FFF2CC'})

    status_to_format = {
        # No grading by default is blank
        GradingStatus.GRADED_POOR: red_format,
        GradingStatus.GRADED_LATE: yellow_format,
        GradingStatus.GRADED: green_format,
        GradingStatus.GRADED_EXCEPTIONAL: blue_format
    }

    # Shuffle all members who have 2 members
    repositories = shuffle_until_no_two_members(repositories)
    teams_with_less_than_two = 0

    for i, data in enumerate(repositories):
        # Grab the team name and member count
        if not data.get_team():
            worksheet.write(
                get_cell_index(ColumnName.TEAM_NAME, (i + 2)),
                "NO TEAM"
            )
        else:
            worksheet.write(
                get_cell_index(ColumnName.TEAM_NAME, (i + 2)),
                data.get_team().name
            )

        worksheet.write(
            get_cell_index(ColumnName.GITHUB_LINK, (i + 2)),
            data.get_repository().html_url
        )

        # Write a comment if the team does not have 2 members
        if data.get_member_count() != 2:
            worksheet.write(
                get_cell_index(ColumnName.COMMENT, (i + 2)),
                "Member Count: %s" % data.get_member_count(),
                workbook.add_format({'bg_color': '#E6B8B7'})
            )
            teams_with_less_than_two += 1

        # By default, the grading status is not graded
        worksheet.write(
            get_cell_index(ColumnName.GRADING_STATUS, (i + 2)),
            GradingStatus.NOT_GRADED.value
        )

        # Create a dropdown for the grading status
        worksheet.data_validation(
            get_cell_index(ColumnName.GRADING_STATUS, (i + 2)),
            {'validate': 'list', 'source': [status.value for status in GradingStatus]}
        )

        # Apply Conditional Formatting to the entire row, based on the value of column F
        for status, format in status_to_format.items():
            worksheet.conditional_format('A%s:G%s' % ((i + 2), (i + 2)),
                                         {
                                             'type': 'formula',
                                             'criteria': '=$F$%d="%s"' % ((i + 2), status.value),
                                             'format': format
                                         })

    # Assign TAs
    instructors = {
        "LAB_INSTRUCTORS": [
            "Jann Garcia",
            "Jose Ortiz",
            "Jose Cordero",
            "Robdiel Melendez",
            "Jomard Concepcion",
            "Misael Mercado",
        ],
        "GRADERS": [
            "Eithan Capella",
            "Christian Perez"
        ]
    }

    # THIS LINE IS NOT REALLY NECESSARY BUT I WANT TO BE SUPER FAIR
    # Shuffle the grader list so that the first graders aren't the same
    # everytime when assigning the very last repositories.
    shuffle(instructors["GRADERS"])

    # Calculate TA/GRADER split to 60/40 ratio respectively
    valid_repos = len(repositories) - teams_with_less_than_two
    ratio_lab = 0.60

    # Number of projects per group
    total_lab_projects = floor(ratio_lab * valid_repos)
    total_grader_projects = valid_repos - total_lab_projects

    # Projects per lab TA and grader
    projects_per_ta = total_lab_projects // len(instructors["LAB_INSTRUCTORS"])

    # Distribute projects to lab TAs
    repo_idx = 2
    for ta in instructors["LAB_INSTRUCTORS"]:
        for _ in range(projects_per_ta):
            worksheet.write(
                get_cell_index(ColumnName.TA, repo_idx),
                ta
            )
            repo_idx += 1

    # Distribute remaining projects to graders in a round-robin fashion
    grader_index = 0
    for _ in range(total_grader_projects):
        grader = instructors["GRADERS"][grader_index % len(instructors["GRADERS"])]
        worksheet.write(
            get_cell_index(ColumnName.TA, repo_idx),
            grader
        )
        repo_idx += 1
        grader_index += 1


    # Autofit the column widths, and save the file
    worksheet.autofit()
    workbook.close()

    print("Successfully created the excel file")


if __name__ == "__main__":
    main()
