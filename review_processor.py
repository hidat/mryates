import os
import argparse
import operator
from review_parser import parser


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
                    parser.DocParser.processFile(file_name, self.albums)

    def exportAlbums(self, targetFilename):
        file = open(targetFilename, "w", encoding='utf-8')
        sortedAlbums = sorted(self.albums, key=operator.attrgetter('filename', 'rotation', 'name'))
        header = "%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\n" % ('Current Rotation', 'Review Title', 'Review', 'Dalet Review', 'Reviewed By', 'One Star', 'Two Star', 'Three Star', 'Filename', 'Artist Credit', 'Label')
        file.write(header)
        for a in sortedAlbums:
            s = a.formatCSV()
            file.write(s)
            file.write('\n')
        file.close()

def main():
    argParser = argparse.ArgumentParser(description='Processes a KEXP weekly review documents into a spreadsheet.')
    argParser.add_argument('input_directory', help="Directory or file to process.  Only docx files are will be processed.")
    argParser.add_argument('output_file', help="Name for the tab-delimited file that will be created to store the results in.")
    #argParser.add_argument('-d', '--delete', default=False, const=True, nargs='?', help="Delete audio files from input_directory after processing")

    args = argParser.parse_args()

    if os.path.isfile(args.input_directory):
        fp = parser.DocParser(args.input_directory)
        fp.process()
        fp.exportAlbums(args.output_file)
    else:
        dp = DirectoryProcessor(args.input_directory)
        dp.processDirectory()
        dp.exportAlbums(args.output_file)


if __name__ == "__main__":
    main()

