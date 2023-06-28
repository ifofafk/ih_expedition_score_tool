class PlayerScore:

    def __init__(self, rank, name, score):
        self.rank = rank
        self.name = name
        self.score = score

    # 实现tostring
    def __str__(self):
        return '名次：%s 名称：%s 分数:%s' % (self.rank, self.name, self.score)

    def get_rank(self):
        return self.rank

    def get_name(self):
        return self.name

    def get_score(self):
        return self.score
