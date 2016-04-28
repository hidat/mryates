import mammoth
import re
import os
import argparse


class Track:
    reNumberExtract = re.compile(r"(\d+)")

    def __init__(self, rawTrack):
        self.rawTrack = rawTrack
        self.stars = 1
        m = self.reNumberExtract.search(rawTrack)

        if m:
            self.trackNum = m.group(0)
        else:
            self.trackNum = None

        if "<em>" in rawTrack:
            self.stars +=1
        if "<strong>" in rawTrack:
            self.stars += 1

class Album:
    reTrackSplit = re.compile(r",|&amp;")

    def __init__(self, filename, rotation):
        self.filename = filename
        self.rotation = rotation
        self.name = None
        self.artistCredit = None
        self.label = None
        self.review = None
        self.trackList = None
        self.tracks = None
        self.reviewedBy = None
        self.oneStarTracks = []
        self.twoStarTracks = []
        self.threeStarTracks = []

    @staticmethod
    def isNameString(nameStr):
        # Checks if the string is in the format of [ALBUM] - [ARTIST CREDIT] ([LABEL])
        labelStart = nameStr.rfind('(')
        if labelStart <= 0:
            return False

        # First try and split by the funky unicode 'en dash' that word likes to use
        parts = nameStr[:(labelStart-1)].split(' \u2013 ')
        if len(parts) == 1:
            # Try splitting with ascii hyphen
            parts = nameStr[:(labelStart-1)].split(' - ')

        if len(parts) == 1:
            return False
        return True

    def parseNameString(self, nameStr):
        # Formatted as [ALBUM] - [ARTIST CREDIT] ([LABEL])
        nameStr = nameStr.replace("&amp;", '&')
        labelStart = nameStr.rfind('(')
        if labelStart > 0:
            self.label = nameStr[labelStart+1:-1]
        # First try and split by the funky unicode 'en dash' that word likes to use
        parts = nameStr[:(labelStart-1)].split(' \u2013 ')
        if len(parts) == 1:
            # Try splitting with ascii hyphen
            parts = nameStr[:(labelStart-1)].split(' - ')

        if len(parts) == 1:
            self.name = parts[0]
        elif len(parts) > 1:
            self.artistCredit = parts[0].strip()
            self.name = parts[1].strip()
            if len(parts) > 2:
                self.name += ' - ' + parts[2]


    def parseReviewString(self, reviewStr):
        self.review = reviewStr
        sentences = reviewStr.split('.')
        if len(sentences) > 2:
            self.reviewedBy = sentences[-1]
            self.trackList = sentences[-2]
            self.tracks = []
            rawTracks = self.reTrackSplit.split(self.trackList)
            for t in rawTracks:
                track = Track(t)
                self.tracks.append(track)
                if track.stars == 3:
                    self.threeStarTracks.append(track)
                elif track.stars == 2:
                    self.twoStarTracks.append(track)
                else:
                    self.oneStarTracks.append(track)

class DirectoryProcessor:
    def __init__(self, directoryPath):
        self.directoryPath = directoryPath
        self.albums = []

    @staticmethod
    def printAlbum(album):
        print("\n%s: %s by [%s] (on %s) - %s" %(album.rotation, album.name, album.artistCredit, album.label, album.review))
        print("Reviewed by: %s.  Tracks to try: %s" % (album.reviewedBy, album.trackList))
        for t in album.tracks:
            print("%s: %d" % (t.trackNum, t.stars))

    def exportAlbum(self, album):
        oneStar = ''
        for t in album.oneStarTracks:
            if not (t.trackNum is None):
                if oneStar > '':
                    oneStar += ', '
                oneStar += str(t.trackNum)

        twoStar = ''
        for t in album.twoStarTracks:
            if not (t.trackNum is None):
                if twoStar > '':
                    twoStar += ', '
                twoStar += str(t.trackNum)

        threeStar = ''
        for t in album.threeStarTracks:
            if not (t.trackNum is None):
                if threeStar > '':
                    threeStar += ', '
                threeStar += str(t.trackNum)

        s = "%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s" % (album.filename, album.rotation, album.artistCredit, album.name, album.label, album.review, oneStar, twoStar, threeStar)
        return s


    def processFile(self, filename):
        style_map = "u => em"

        with open(filename, "rb") as docx_file:
            result = mammoth.convert_to_html(docx_file, style_map=style_map)
            html = result.value # The generated HTML
            paras = html.split('<p>')
            currentRotation = None

            album = None

            for p in paras:
                s = p[:-4].rstrip().replace(u'\xa0', ' ')
                if len(s) > 0:
                    if len(s) < 10:
                        # Check to make sure this isn't a random <br /> or other html
                        if s[0] != '<':
                            currentRotation = s[:-1].rstrip()
                            albumName = None
                            albumReview = None
                            waitingForAlbum = True
                    else:
                        if album is None:
                            if Album.isNameString(s):
                                album = Album(filename, currentRotation)
                                album.parseNameString(s)
                        else:
                            album.parseReviewString(s)
                            self.albums.append(album)
                            album = None

    def processDirectory(self):
        path_start = len(self.directoryPath) + 1
        for root, dir, files in os.walk(self.directoryPath):
            if len(root) > path_start:
                path = root[path_start:]
            else:
                path = ''
            for src_name in files:
                file_name = os.path.join(root, src_name)
                print(src_name)
                self.processFile(file_name)

    def exportAlbums(self, targetFilename):
        file = open(targetFilename, "w", encoding='utf-8')
        for a in self.albums:
            s = self.exportAlbum(a)
            file.write(s)
            file.write('\n')
        file.close()

dp = DirectoryProcessor("reviews_2015")
dp.processDirectory()
dp.exportAlbums('reviews_2015.txt')

