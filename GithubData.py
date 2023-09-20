class GithubData():
    def __init__(self, repository):
        self.repository = repository
        self.team = repository.get_teams()[0]
        self.member_count = self.team.get_members().totalCount

    def get_repository(self):
        return self.repository
    
    def get_team(self):
        return self.team
    
    def get_member_count(self):
        return self.member_count
    