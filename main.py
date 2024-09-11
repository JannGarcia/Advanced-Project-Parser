from math import floor, ceil
from github import Github, Auth
from dotenv import load_dotenv
import os
from enum import Enum
import xlsxwriter
from random import shuffle
from GithubData import GithubData
from compiler import compile_projects
import sys

ORGANIZATION_NAME = "UPRM-CIIC4010-F24"
PROJECT_PREFIX = "pa1"
FILE_NAME = PROJECT_PREFIX + "-Distribution.xlsx"


class ColumnName(Enum):
    TEAM_NAME = "Team Name"
    GITHUB_LINK = "Github Link"
    MEMBER_1 = "Member 1"
    MEMBER_2 = "Member 2"
    TA = "TA"
    GRADING_STATUS = "Grading Status"
    COMMENT = "Comment"
    RELEASE = "Release"
    COMPILES = "Compiles"


class GradingStatus(Enum):
    NOT_GRADED = "Not Graded"
    GRADED_POOR = "Graded (<=50%)"
    GRADED_LATE = "Graded (Late)"
    GRADED = "Graded"
    GRADED_EXCEPTIONAL = "Graded (>=100%)"


column_name_to_index = {
    ColumnName.TEAM_NAME: "A",
    ColumnName.GITHUB_LINK: "B",
    ColumnName.MEMBER_1: "C",
    ColumnName.MEMBER_2: "D",
    ColumnName.TA: "E",
    ColumnName.GRADING_STATUS: "F",
    ColumnName.COMMENT: "G",
    ColumnName.RELEASE: "H",
}


def get_cell_index(column_name, i):
    return f"{column_name_to_index[column_name]}{i}"


def open_workbook():
    workbook = xlsxwriter.Workbook(FILE_NAME)
    worksheet = workbook.add_worksheet()

    header_cell = workbook.add_format({
        'bold': True,
        'align': 'center',
        'bg_color': '#D3D3D3'
    })

    for column_name, column_index in column_name_to_index.items():
        worksheet.write(f"{column_index}1", column_name.value, header_cell)

    return workbook, worksheet


def get_token():
    load_dotenv()
    token = os.getenv("GITHUB_TOKEN")

    if not isinstance(token, str):
        print(f"Token value expected string, got {type(token)}")
        print(f"Perhaps you have not loaded your env file?")
        exit(400)

    return token


def login(token):
    auth = Auth.Token(token)
    g = Github(auth=auth)
    print("Successfully Logged in as: " + g.get_user().login)
    return g


def get_repositories():
    g = login(get_token())
    organization = g.get_organization(ORGANIZATION_NAME)
    print("Entered Organization: " + organization.login)

    repos = [
        GithubData(repo)
        for repo in organization.get_repos()
        if repo.name.lower().startswith(PROJECT_PREFIX)
        and repo.name.lower() != PROJECT_PREFIX
    ]

    if not repos:
        raise ValueError(
            "No repositories found for this project.\n"
            "Check your token's permissions to allow access."
        )

    return repos


def shuffle_until_no_two_members(repos):
    repositories = sorted(repos, key=lambda repo: repo.get_member_count(), reverse=True)

    reverse_index = 0
    for i, data in enumerate(repositories):
        reverse_index = i
        if data.get_member_count() != 2:
            break

    first_half = repositories[:reverse_index]
    shuffle(first_half)
    repositories = first_half + repositories[reverse_index:]
    return repositories


def assign_tas(workbook, worksheet, repositories, teams_with_less_than_two):
    instructors = {
        "LAB_INSTRUCTORS": [
            "Jann Garcia",
            "Jose Ortiz",
            "Jose Cordero",
            "Robdiel Melendez",
            "Jomard Concepcion",
            "Misael Mercado",
        ],
        "GRADERS": ["Eithan Capella", "Christian Perez"],
    }

    shuffle(instructors["GRADERS"])
    valid_repos = len(repositories) - teams_with_less_than_two
    ratio_lab = 0.60
    total_lab_projects = floor(ratio_lab * valid_repos)
    total_grader_projects = valid_repos - total_lab_projects
    projects_per_ta = total_lab_projects // len(instructors["LAB_INSTRUCTORS"])

    repo_idx = 2
    for ta in instructors["LAB_INSTRUCTORS"]:
        for _ in range(projects_per_ta):
            if repo_idx > valid_repos + 1:
                break
            worksheet.write(get_cell_index(ColumnName.TA, repo_idx), ta)
            repo_idx += 1

    grader_index = 0
    while repo_idx <= valid_repos + 1:
        grader = instructors["GRADERS"][grader_index % len(instructors["GRADERS"])]
        worksheet.write(get_cell_index(ColumnName.TA, repo_idx), grader)
        repo_idx += 1
        grader_index += 1


def write_repository_data(workbook, worksheet, repositories):
    red_format = workbook.add_format({"bg_color": "#F8CECC"})
    green_format = workbook.add_format({"bg_color": "#C6EFCE"})
    blue_format = workbook.add_format({"bg_color": "#BDD7EE"})
    yellow_format = workbook.add_format({"bg_color": "#FFF2CC"})

    status_to_format = {
        GradingStatus.GRADED_POOR: red_format,
        GradingStatus.GRADED_LATE: yellow_format,
        GradingStatus.GRADED: green_format,
        GradingStatus.GRADED_EXCEPTIONAL: blue_format,
    }

    shuffled_repositories = shuffle_until_no_two_members(repositories)
    teams_with_less_than_two = 0

    for i, data in enumerate(shuffled_repositories, start=2):
        team = data.get_team()
        worksheet.write(
            get_cell_index(ColumnName.TEAM_NAME, i),
            "NO TEAM" if not team else team.name,
        )
        worksheet.write(
            get_cell_index(ColumnName.GITHUB_LINK, i),
            data.get_repository().html_url,
        )

        if data.get_member_count() != 2:
            worksheet.write(
                get_cell_index(ColumnName.COMMENT, i),
                f"Member Count: {data.get_member_count()}",
                workbook.add_format({"bg_color": "#E6B8B7", "text_wrap": True}),
            )
            teams_with_less_than_two += 1

        release_status = "Release Found" if data.has_release() else "No Release Found"
        worksheet.write(
            get_cell_index(ColumnName.RELEASE, i),
            release_status,
            workbook.add_format(
                {"bg_color": "#C6EFCE"}
                if data.has_release()
                else {"bg_color": "#E6B8B7"}
            ),
        )

        worksheet.write(
            get_cell_index(ColumnName.GRADING_STATUS, i),
            GradingStatus.NOT_GRADED.value,
        )
        worksheet.data_validation(
            get_cell_index(ColumnName.GRADING_STATUS, i),
            {"validate": "list", "source": [status.value for status in GradingStatus]},
        )

        for status, fmt in status_to_format.items():
            worksheet.conditional_format(
                f"A{i}:G{i}",
                {
                    "type": "formula",
                    "criteria": f'=$F${i}="{status.value}"',
                    "format": fmt,
                },
            )

    return teams_with_less_than_two, shuffled_repositories


def main():

    if "compile" in sys.argv:
        # Add the COMPILES column
        column_name_to_index[ColumnName.COMPILES] = "I"

    repositories = get_repositories()
    workbook, worksheet = open_workbook()
    teams_with_less_than_two, shuffled_repositories = write_repository_data(workbook, worksheet, repositories)
    assign_tas(workbook, worksheet, shuffled_repositories, teams_with_less_than_two)

    # Optional: If the user passes in a flag called "compile", then compile the projects and log the results
    if "compile" in sys.argv:
        print("Compiling projects...")
        repo_urls = [repo.get_url() for repo in shuffled_repositories]
        results = compile_projects(repo_urls)

        # Create a mapping from URL to compilation log for faster access
        url_to_log = {url: logs for url, logs in results.items()}

        for i, repo in enumerate(shuffled_repositories, start=2):
            url = repo.get_url()
            logs = url_to_log.get(url, [])
            compilation_log = logs[1] if len(logs) > 1 else ""

            if "compiled successfully" in compilation_log.lower():
                worksheet.write(
                    get_cell_index(ColumnName.COMPILES, i),
                    "Compiled Successfully",
                    workbook.add_format({"bg_color": "#C6EFCE"})
                )
            else:
                print(f"Error compiling project {url}: {compilation_log}")
                worksheet.write(
                    get_cell_index(ColumnName.COMPILES, i),
                    "Compilation Error",
                    workbook.add_format({"bg_color": "#E6B8B7"})
                )

    worksheet.autofit()
    workbook.close()
    print("Successfully created the excel file")


if __name__ == "__main__":
    main()
