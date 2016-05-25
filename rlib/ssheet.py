import smartsheet

class SSRelease:
    def __init__(self, cells):
        self.title = cells[3].value
        self.artist = cells[2].value
        self.rotation = cells[1].value
        self.obscenityRating = cells[6].value
        self.mbID = cells[9].value

class SDK:
    def __init__(self, apiKey):
        self.apiKey = apiKey
        self.sdk = smartsheet.Smartsheet(apiKey)

    def readWeeklySheet(self, sheetId):
        releases = []
        sheet = self.sdk.Sheets.get_sheet(sheetId)
        for r in sheet.rows:
            releases.append(SSRelease(r.cells))
        return releases
