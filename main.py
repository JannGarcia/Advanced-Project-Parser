from github import Github, Auth
from dotenv import load_dotenv
import os # Used to access the .env file

import xlsxwriter
import asyncio

ORGANAZATION_NAME = "UPRM-CIIC4010-F23"
PROJECT_PREFIX = "pa0"

def open_workbook():
    workbook = xlsxwriter.Workbook('%s-Distribution.xlsx' % PROJECT_PREFIX)
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
    worksheet.write('E1', 'Comment', header_cell)

    return (workbook, worksheet)

def get_repositories():
    g = login(get_token())

    organization = g.get_organization(ORGANAZATION_NAME)
    print("Entered Organization: " + organization.login)

    return [
        repo for repo in organization.get_repos() 
        if repo.name.lower().startswith(PROJECT_PREFIX) 
        and repo.name.lower() != PROJECT_PREFIX
    ]

def get_token():
    # Load the .env file
    load_dotenv()
    return os.getenv("GITHUB_TOKEN")

def login(token):
    auth = Auth.Token(token)
    g = Github(auth=auth)
    print("Successfully Logged in as: " + g.get_user().login)
    return g

def main(): 
    repositories = get_repositories()
    workbook, worksheet = open_workbook()

    for i, repo in enumerate(repositories):
        # Grab the team name 
        team = repo.get_teams()[0]
        member_count = team.get_members().totalCount
        worksheet.write('A%s' % (i+2), team.name)
        worksheet.write('B%s' % (i+2), repo.html_url)

        # Write a comment if the team does not have 2 members
        if member_count != 2:
            worksheet.write('E%s' % (i+2), "Member Count: %s" % member_count, workbook.add_format({'bg_color': '#E6B8B7'}))

    worksheet.autofit()
    workbook.close()

    print("Successfully created the excel file")

if __name__ == "__main__":
    main()