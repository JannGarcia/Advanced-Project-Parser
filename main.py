from github import Github, Auth
from dotenv import load_dotenv
import os # Used to access the .env file

ORGANAZATION_NAME = "UPRM-CIIC4010-F23"
PROJECT_PREFIX = "pa0"

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
    g = login(get_token())

    organization = g.get_organization(ORGANAZATION_NAME)
    print("Organization: " + organization.login)

if __name__ == "__main__":
    main()