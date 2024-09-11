class GithubData:
    def __init__(self, repository):
        self.repository = repository

        teams = repository.get_teams()
        if teams.totalCount > 0:
            self.team = repository.get_teams()[0]
            self.member_count = self.team.get_members().totalCount

        else:
            self.team = None
            self.member_count = 0

    def get_repository(self):
        return self.repository

    def get_team(self):
        return self.team

    def get_member_count(self):
        return self.member_count

    def get_url(self):
        return self.repository.html_url

    def has_release(self):
        return self.repository.get_releases().totalCount > 0
