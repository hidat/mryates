__author__ = 'hidat'

import os
import argparse
from rlib import parser, serializer

def main():
    argParser = argparse.ArgumentParser(description='Processes a KEXP weekly review documents into a spreadsheet.')
    argParser.add_argument('input_file', help="Directory or file to process.  Only docx files are will be processed.")
    argParser.add_argument('output_directory', help="Name for the tab-delimited file that will be created to store the results in.")
    #argParser.add_argument('-d', '--delete', default=False, const=True, nargs='?', help="Delete audio files from input_file after processing")

    args = argParser.parse_args()

    if os.path.isfile(args.input_file):
        successCount = 0
        errorCount = 0
        reviews = parser.CSVFile.readFile(args.input_file)
        fp = serializer.DaletSerializer(args.output_directory)
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
