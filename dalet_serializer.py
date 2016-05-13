__author__ = 'hidat'

from yattag import Doc, indent
import os.path as path
import os
import argparse
import csv

class DaletSerializer:

    def __init__(self, output_dir):
        # Set up metadata directories
        self.release_meta_dir = output_dir
        if not path.exists(self.release_meta_dir):
            os.makedirs(self.release_meta_dir)
        print("Release meta: ", self.release_meta_dir)


    def saveRelease(self, release):
        """
        Create an XML file of release metadata that Dalet will be happy with
        
        :param release: Processed release metadata from MusicBrainz
        """

        output_dir = self.release_meta_dir

        doc, tag, text = Doc().tagtext()

        doc.asis('<?xml version="1.0" encoding="UTF-8"?>')
        with tag('Titles'):
            with tag('GlossaryValue'):
                with tag('Key1'):
                    text(release.mbID)
                with tag('ItemCode'):
                    text(release.mbID)
                with tag('KEXPReviewRich'):
                    text(release.daletReview)
        formatted_data = indent(doc.getvalue())

        output_file = path.join(output_dir, 'r' + release.mbID + ".xml")
        with open(output_file, "wb") as f:
            f.write(formatted_data.encode("UTF-8"))

class Review:

    def __init__(self, csvParts):
        if csvParts is not None:
            if len(csvParts) > 1:
                #header = "%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\n" % ('Current Rotation', 'Review Title', 'Review', 'Dalet Review', 'Reviewed By', 'One Star', 'Two Star', 'Three Star', 'Filename', 'Artist Credit', 'Label')
                self.rotation = csvParts[0]
                self.name = csvParts[1]
                self.review = csvParts[2]
                self.daletReview = csvParts[3]
                self.reviewedBy = csvParts[4]
                self.filename = csvParts[8]
                self.artistCredit = csvParts[9]
                self.label = csvParts[10]
                self.mbID = csvParts[11].strip()
                self.oneStarTracks = csvParts[5].split(',')
                self.twoStarTracks = csvParts[6].split(',')
                self.threeStarTracks = csvParts[7].split(',')

        else:
            self.filename = None
            self.rotation = None
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
            self.mbID = None

    @staticmethod
    def readCSV(filename):
        reviews = []
        f = open(filename, 'r')
        csvReader = csv.reader(open(filename, newline=''), delimiter='\t', quotechar='"')
        firstLine = True
        for line in csvReader:
            if not firstLine:
                review = Review(line)
                if review is not None:
                    reviews.append(review)
            else:
                firstLine = False

        return reviews



def main():
    parser = argparse.ArgumentParser(description='Processes a KEXP weekly review documents into a spreadsheet.')
    parser.add_argument('input_file', help="Directory or file to process.  Only docx files are will be processed.")
    parser.add_argument('output_directory', help="Name for the tab-delimited file that will be created to store the results in.")
    #parser.add_argument('-d', '--delete', default=False, const=True, nargs='?', help="Delete audio files from input_file after processing")

    args = parser.parse_args()

    if os.path.isfile(args.input_file):
        successCount = 0
        errorCount = 0
        reviews = Review.readCSV(args.input_file)
        fp = DaletSerializer(args.output_directory)
        for review in reviews:
            if review.mbID is not None and review.mbID > '':
                fp.saveRelease(review)
                successCount+=1
            else:
                errorCount+=1
        print("Done processing file, %d reviews exported, %d reviews did not have MBID's" % (successCount, errorCount))
    else:
        print(args.input_file + " file not found!")


if __name__ == "__main__":
    main()
