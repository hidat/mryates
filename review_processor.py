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
        if len(nameStr) > 200:
            return False

        # Checks if the string is in the format of [ALBUM] - [ARTIST CREDIT] ([LABEL])
        labelStart = nameStr.rfind('(')
        if labelStart <= 0:
            return False

        s = nameStr[:(labelStart-1)]
        parts = s.split('- ')
        if len(parts) > 1:
            return True

        parts = s.split(' -')
        if len(parts) > 1:
            return True

        parts = s.split(': ')
        if len(parts) > 1:
            return True

        return False

    def parseNameString(self, nameStr):
        # Formatted as [ALBUM] - [ARTIST CREDIT] ([LABEL])
        nameStr = nameStr.replace("&amp;", '&')
        labelStart = nameStr.rfind('(')
        if labelStart > 0:
            self.label = nameStr[labelStart+1:-1]

        parts = nameStr[:(labelStart-1)].split('- ')
        if len(parts) == 1:
            parts = nameStr[:(labelStart-1)].split(' -')
        if len(parts) == 1:
            parts = nameStr[:(labelStart-1)].split(': ')

        if len(parts) == 1:
            self.name = parts[0]
        elif len(parts) > 1:
            self.artistCredit = parts[0].strip()
            self.name = parts[1].strip()
            if len(parts) > 2:
                self.name += ' - ' + parts[2]

    ###
    # Parses a paragraph that we are expecting to be a 'review'
    # This can be called multiple times (because sometimes we get newlines in the middle of a review).  Multiples lines
    # are appended together.
    ###
    def parseReviewString(self, reviewStr):
        if self.review is None:
            self.review = reviewStr
        else:
            self.review += ' ' + reviewStr

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
            lastAlbum = None

            for p in paras:
                s = p[:-4].rstrip().replace(u'\xa0', ' ').replace(u'\u2013', '-')
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
                                lastAlbum = None
                            elif not (lastAlbum is None):
                                # did somebody put a newline in the middle of a review? Try to add it to the last album
                                lastAlbum.parseReviewString(s)
                        else:
                            album.parseReviewString(s)
                            self.albums.append(album)
                            lastAlbum = album
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

def main():
    parser = argparse.ArgumentParser(description='Processes a KEXP weekly review documents into a spreadsheet.')
    parser.add_argument('input_directory', help="Directory containing one or more docx files to process.")
    parser.add_argument('output_file', help="Name of the tab-delimited file to store the results in.")
    #parser.add_argument('-d', '--delete', default=False, const=True, nargs='?', help="Delete audio files from input_directory after processing")

    args = parser.parse_args()

    dp = DirectoryProcessor(args.input_directory)
    dp.processDirectory()
    dp.exportAlbums(args.output_file)


if __name__ == "__main__":
    main()

