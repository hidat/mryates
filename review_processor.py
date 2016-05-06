import mammoth
import re
import os
import argparse
import unicodedata
import html
import operator


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
    reReviewer = re.compile(r"[-]+([\w\s]+)")

    def __init__(self, filename, rotation):
        self.filename = filename
        self.rotation = rotation
        self.name = None
        self.artistCredit = None
        self.label = None
        self.review = None
        self.trackList = None
        self.tracks = None
        self.reviewedBy = ''
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
        if s.find('- ') > -1:
            return True

        if s.find(' -') > -1:
            return True

        if s.find(': ') > -1:
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
        tl = None
        rb = None
        if len(sentences) > 1:
            rb = sentences[-1].strip()
            if rb.startswith('-'):
                rm = self.reReviewer.findall(rb)
                if len(rm) > 0:
                    self.reviewedBy = rm[0]
            elif rb.startswith('Try'):
                parts = rb.split('-')
                if len(parts) > 1:
                    tl = parts[0]
                    rm = self.reReviewer.findall(rb)
                    if len(rm) > 0:
                        self.reviewedBy = rm[0]
            else:
                rb = None

        if tl is None and len(sentences) > 2:
            tl = sentences[-2].strip()
        if (not (tl is None)) and tl.startswith('Try'):
                self.trackList = tl

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

    def print(self):
        print("\n%s: %s by [%s] (on %s) - %s" %(self.rotation, self.name, self.artistCredit, self.label, self.review))
        print("Reviewed by: %s.  Tracks to try: %s" % (self.reviewedBy, self.trackList))
        for t in self.tracks:
            print("%s: %d" % (t.trackNum, t.stars))

    def formatCSV(self):
        oneStar = ''
        for t in self.oneStarTracks:
            if not (t.trackNum is None):
                if oneStar > '':
                    oneStar += ', '
                oneStar += str(t.trackNum)

        twoStar = ''
        for t in self.twoStarTracks:
            if not (t.trackNum is None):
                if twoStar > '':
                    twoStar += ', '
                twoStar += str(t.trackNum)

        threeStar = ''
        for t in self.threeStarTracks:
            if not (t.trackNum is None):
                if threeStar > '':
                    threeStar += ', '
                threeStar += str(t.trackNum)

        decodedReview = html.unescape(self.review)
        textReview = re.sub("<.*?>", "", decodedReview)
        decodedReview = decodedReview.replace("em>", "u>").replace("strong>", "b>")
        encodedReview = html.escape(decodedReview, False)
        s = "%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s" % (self.rotation, self.name, textReview, encodedReview, self.reviewedBy, oneStar, twoStar, threeStar, self.filename, self.artistCredit, self.label)
        return s
    
class FileProcessor:
    def __init__(self, filename):
        self.filename = filename
        self.albums = []

    def process(self):
        FileProcessor.processFile(self.filename, self.albums)

    def exportAlbums(self, targetFilename):
        file = open(targetFilename, "w", encoding='utf-8')
        sortedAlbums = sorted(self.albums, key=operator.attrgetter('filename', 'rotation', 'name'))
        for a in sortedAlbums:
            s = a.formatCSV()
            file.write(s)
            file.write('\n')
        file.close()

    @staticmethod
    def processFile(filename, albums):
        style_map = "u => em"

        with open(filename, "rb") as docx_file:
            result = mammoth.convert_to_html(docx_file, style_map=style_map)
            html = result.value # The generated HTML
            paras = html.split('<p>')
            currentRotation = None

            album = None
            lastAlbum = None

            for p in paras:
                s = p[:-4].strip()
                #s = s.replace(u'\xa0', ' ').replace(u'\u2013', '-')
                s = unicodedata.normalize('NFKC', s).replace(u'\u2013', '-')
                if len(s) > 0 and s[0] != '<':
                    if len(s) > 0:
                        if len(s) < 10:
                            s = s[:-1].rstrip()
                            if s=='H' or s=='M' or s=='L' or s=='R/N':
                                currentRotation = s
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
                                albums.append(album)
                                lastAlbum = album
                                album = None


class DirectoryProcessor:
    def __init__(self, directoryPath):
        self.directoryPath = directoryPath
        self.albums = []


    def processDirectory(self):
        path_start = len(self.directoryPath) + 1
        for root, dir, files in os.walk(self.directoryPath):
            if len(root) > path_start:
                path = root[path_start:]
            else:
                path = ''
            for src_name in files:
                if src_name[-4:].lower() == 'docx' and src_name[:1] != '~':
                    file_name = os.path.join(root, src_name)
                    print(src_name)
                    FileProcessor.processFile(file_name, self.albums)

    def exportAlbums(self, targetFilename):
        file = open(targetFilename, "w", encoding='utf-8')
        sortedAlbums = sorted(self.albums, key=operator.attrgetter('filename', 'rotation', 'name'))
        for a in sortedAlbums:
            s = a.formatCSV()
            file.write(s)
            file.write('\n')
        file.close()

def main():
    parser = argparse.ArgumentParser(description='Processes a KEXP weekly review documents into a spreadsheet.')
    parser.add_argument('input_directory', help="Directory or file to process.  Only docx files are will be processed.")
    parser.add_argument('output_file', help="Name for the tab-delimited file that will be created to store the results in.")
    #parser.add_argument('-d', '--delete', default=False, const=True, nargs='?', help="Delete audio files from input_directory after processing")

    args = parser.parse_args()

    if os.path.isfile(args.input_directory):
        fp = FileProcessor(args.input_directory)
        fp.process()
        fp.exportAlbums(args.output_file)
    else:
        dp = DirectoryProcessor(args.input_directory)
        dp.processDirectory()
        dp.exportAlbums(args.output_file)


if __name__ == "__main__":
    main()

