import os
import argparse
from rlib import parser, ssheet, serializer

def findRelease(review,  releases):
    # First try to find by title
    titleToFind = review.name.lower()
    for release in releases:
        if release.title.lower() == titleToFind:
            return release
    return None

def mergeReviewsAndReleases(reviews, releases):
    for review in reviews:
        r = findRelease(review, releases)
        if r is not None:
            review.mbID = r.mbID
    return reviews

def getMbIdFromUser(review):
    mbID = input("Please enter the MBID for '%s' by %s: " % (review.name, review.artistCredit))
    return mbID

def printReviews(reviews):
    foundCount = 0
    missingCount = 0
    missing = []
    for review in reviews:
        mbID = review.mbID
        if mbID is None:
            missingCount += 1
            missing.append(review)
        else:
            foundCount += 1
            print("'%s' by %s: %s" % (review.name, review.artistCredit, mbID))

    if missingCount > 0:
        print('\nReviews missing MusicBrainz IDs:')
        for review in missing:
            print("    '%s' by %s" % (review.name, review.artistCredit))

    return foundCount

def fillInMissingReviews(reviews):
    missingCount = 0
    for review in reviews:
        if review.mbID is None:
            mbID = getMbIdFromUser(review)
            if mbID is not None and mbID > '':
                review.mbID = mbID
            else:
                missingCount += 1
    return missingCount

def exportReviews(reviews, outputDirectory):
    errorCount = 0
    fp = serializer.DaletSerializer(outputDirectory)
    for review in reviews:
        if review.mbID is not None and review.mbID > '':
            fp.saveRelease(review)
        else:
            errorCount+=1
    return errorCount


def main():
    argParser = argparse.ArgumentParser(description='Processes a KEXP weekly review documents and generate Dalet review import.')
    argParser.add_argument('input_file', help="Word document to process.  Only docx files are supported.")
    argParser.add_argument('api_key', help="Your Smartsheet API Key")
    argParser.add_argument('-d', '--directory', default="dalet", help="Directory to put Dalet Impex files in.")
    argParser.add_argument('-w', '--worksheet', help="Smartsheet Worksheet ID that contains the reviews associated MusicBrainz ID's")

    args = argParser.parse_args()

    if os.path.isfile(args.input_file):
        if args.worksheet is None:
            sheetId = '3416095957772164'
        else:
            sheetId = args.worksheet

        sdk = ssheet.SDK(args.api_key)
        fp = parser.DocParser(args.input_file)
        reviews = fp.process()
        releases = sdk.readWeeklySheet(sheetId)
        mergedReviews = mergeReviewsAndReleases(reviews, releases)
        foundCount = printReviews(mergedReviews)
        missingCount = len(mergedReviews) - foundCount
        if missingCount > 0:
            ans = input("Would you like to enter missing MusicBrainz IDs now? [Y]n: ")
            if ans.lower() == 'y' or ans == '':
                missingCount = fillInMissingReviews(reviews)
        doExport = missingCount == 0
        if missingCount > 0:
            ans = input("There are still %s reviews missing MusicBrainz IDs, would you like to export to Dalet amyway? y[N]: "  % (missingCount))
            if ans.lower() == 'y':
                doExport = True
        if doExport:
            exportReviews(mergedReviews, args.directory)
    else:
        print(args.input_file + " does not appear to be a valid file.")

if __name__ == "__main__":
    main()
