class sqlfilehandler(object):
    def __init__(self, sqlfilename):
        file = open(sqlfilename, 'r')
        self.content = file.read()
        data = self.content.split('#')
        self.queries = []
        self.sheet = []
        for query in data:
            if '--sheet' in query:
                query = query.split('--sheet')
                self.sheet.append(query[0].strip())
                self.queries.append(query[1].strip())

    def getSheetnames(self):
        return self.sheet
    
    def getQueries(self):
        return self.queries