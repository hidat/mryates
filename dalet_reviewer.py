import os
import argparse
from rlib import parser, ssheet

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

def main():
    argParser = argparse.ArgumentParser(description='Processes a KEXP weekly review documents and generate Dalet review import.')
    argParser.add_argument('input_file', help="Word document to process.  Only docx files are supported.")
    argParser.add_argument('-k', '--key', default=False, const=True, nargs='?', help="Your Smartsheet API Key")
    argParser.add_argument('-w', '--worksheet', default=False, const=True, nargs='?', help="Smartsheet Worksheet ID that contains the reviews associated MusicBrainz ID's")

    args = argParser.parse_args()

    if os.path.isfile(args.input_file):
        key = args.key
        if key is not None:
            if args.worksheet is None:
                sheetId = '3416095957772164'
            else:
                sheetId = args.worksheet

            sdk = ssheet.SDK(key)
            fp = parser.DocParser(args.input_file)
            reviews = fp.process()
            releases = sdk.readWeeklySheet(sheetId)
            mergedReviews = mergeReviewsAndReleases(reviews, releases)

            foundCount = 0
            for review in mergedReviews:
                if review.mbID is not None:
                    foundCount += 1
                print("%s by %s: %s" % (review.name, review.artistCredit, review.mbID))
            print("\nReleases for %d of %d reviews found." % (foundCount, len(mergedReviews)))
        else:
            print('Please specify your API key')
    else:
        print(args.input_file + " does not appear to be a valid file.")

if __name__ == "__main__":
    main()
